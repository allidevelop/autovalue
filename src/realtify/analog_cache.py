"""Кеш аналогів за адресою/будинком (повторювані адреси клієнта).

Якщо для адреси (на рівні будинку/ЖК) аналоги вже підбиралися — перевикористовуємо їх,
пропускаючи дорогий і крихкий пошук/збір. Сховище: SQLite + копії скриншотів у
``data/analog_cache/``. Ключ будується по місту+вулиці+будинку+типу; номер квартири
відкидається, бо аналоги спільні для будинку.
"""

from __future__ import annotations

import json
import re
import shutil
import sqlite3
from datetime import datetime
from pathlib import Path

from realtify.models import Comparable
from realtify.paths import PROJECT_ROOT

CACHE_ROOT = PROJECT_ROOT / "data" / "analog_cache"
DB_PATH = CACHE_ROOT / "cache.db"

_IMG_FIELDS = ("screenshot_path", "report_image_path")


def _connect() -> sqlite3.Connection:
    CACHE_ROOT.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS analog_cache (
            address_key   TEXT PRIMARY KEY,
            city          TEXT,
            address       TEXT,
            property_type TEXT,
            complex_name  TEXT,
            candidates_json TEXT NOT NULL,
            count         INTEGER NOT NULL DEFAULT 0,
            created_at    TEXT,
            updated_at    TEXT,
            hit_count     INTEGER NOT NULL DEFAULT 0
        )
        """
    )
    return conn


# Тип вулиці → канонічна форма. Витяг з Реєстру пише повну форму («вулиця»,
# «набережна»), таблиця аналога — скорочення («вул.», «наб.»). Без канонізації
# ключі дому не збігаються → building-тир порожній → пошук розповзається на місто.
# Різні типи лишаються РІЗНИМИ («бул.»≠«пл.»), щоб не злити різні вулиці з однаковою назвою.
_STREET_TYPE_CANON: tuple[tuple[str, str], ...] = (
    (r"вулиц[яіеї]|вул", "вул"),
    (r"проспект|просп|пр-кт|пр-т", "просп"),
    (r"бульвар|бульв|б-р", "бул"),
    (r"провулок|пров", "пров"),
    (r"набережн\w*|наб", "наб"),
    (r"майдан", "майдан"),
    (r"площ[аіі]|пл", "пл"),
    (r"узвіз|узв", "узвіз"),
    (r"проїзд", "проїзд"),
    (r"шосе", "шосе"),
)


def address_key(
    *,
    city: str | None,
    address: str | None,
    property_type: str | None,
    complex_name: str | None = None,
) -> str:
    """Стабільний ключ адреси на рівні будинку (без номера квартири).

    Стійкий до варіацій написання: «17-К»=«17К», «будинок/м.»-префікси,
    тип вулиці («вулиця»=«вул.»), дубль міста в city та address.
    Так бібліотека імпорту й оцінка дають один ключ.
    """
    parts = [str(p) for p in (city, address, complex_name) if p]
    raw = " ".join(parts).lower().replace("’", "'")
    # номер квартири/апартаментів відкидаємо — аналоги спільні для будинку
    raw = re.sub(r"\b(?:кв\.?|квартира|апартаменти?|apt\.?)\s*№?\s*\d+\S*", " ", raw)
    # generic-слова будинку/міста
    raw = re.sub(r"\b(?:будинок|буд|місто|м)\b\.?", " ", raw)
    # тип вулиці → канонічна форма (вулиця/вул./просп./наб. …) — до зняття крапок/дефісів
    for _pat, _canon in _STREET_TYPE_CANON:
        raw = re.sub(rf"\b(?:{_pat})\.?(?=\W|$)", f" {_canon} ", raw)
    # усе, крім букв/цифр/пробілів → пробіл (зокрема дефіси, коми, крапки)
    raw = re.sub(r"[^0-9a-zа-яіїєґ'\s]", " ", raw)
    # склеюємо «17 к» → «17к»
    raw = re.sub(r"(\d+)\s+([a-zа-яіїєґ])(?=\s|$)", r"\1\2", raw)
    # дедуп підряд однакових токенів (місто могло потрапити і в city, і в address)
    tokens: list[str] = []
    for token in raw.split():
        if not tokens or tokens[-1] != token:
            tokens.append(token)
    key = "-".join(tokens)
    if property_type:
        key = f"{key}-{str(property_type).strip().lower()}"
    return key.strip("-")[:200]


def lookup(key: str, job_output_dir: Path) -> list[Comparable] | None:
    """Кешовані аналоги (зі скопійованими у папку завдання скриншотами) або None."""
    if not key:
        return None
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT candidates_json FROM analog_cache WHERE address_key = ?", (key,)
        ).fetchone()
        if not row:
            return None
        conn.execute(
            "UPDATE analog_cache SET hit_count = hit_count + 1, updated_at = ? "
            "WHERE address_key = ?",
            (_now(), key),
        )
        conn.commit()
    finally:
        conn.close()

    payload = json.loads(row[0])
    key_dir = CACHE_ROOT / key
    dirs = {
        "screenshot_path": job_output_dir / "screenshots",
        "report_image_path": job_output_dir / "report_images",
    }
    for directory in dirs.values():
        directory.mkdir(parents=True, exist_ok=True)

    result: list[Comparable] = []
    for item in payload:
        data = dict(item)
        for field in _IMG_FIELDS:
            name = data.get(field)
            if not name:
                data[field] = None
                continue
            src = key_dir / Path(str(name)).name
            if src.exists():
                target = dirs[field] / src.name
                shutil.copy2(src, target)
                data[field] = str(target)
            else:
                data[field] = None
        result.append(Comparable.model_validate(data))
    return result or None


def save(
    key: str,
    *,
    city: str | None,
    address: str | None,
    property_type: str | None,
    complex_name: str | None,
    candidates: list[Comparable],
) -> None:
    """Зберігає підібрані аналоги (з копіями скриншотів) під ключем адреси."""
    if not key or not candidates:
        return
    key_dir = CACHE_ROOT / key
    key_dir.mkdir(parents=True, exist_ok=True)

    payload: list[dict] = []
    for candidate in candidates:
        data = json.loads(candidate.model_dump_json())
        for field in _IMG_FIELDS:
            value = data.get(field)
            if not value:
                data[field] = None
                continue
            src = Path(str(value))
            if src.exists():
                try:
                    shutil.copy2(src, key_dir / src.name)
                    data[field] = src.name  # зберігаємо лише ім'я; lookup бере з key_dir
                except OSError:
                    data[field] = None
            else:
                data[field] = None
        payload.append(data)

    now = _now()
    conn = _connect()
    try:
        conn.execute(
            """
            INSERT INTO analog_cache
                (address_key, city, address, property_type, complex_name,
                 candidates_json, count, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(address_key) DO UPDATE SET
                candidates_json = excluded.candidates_json,
                count           = excluded.count,
                complex_name    = excluded.complex_name,
                updated_at      = excluded.updated_at
            """,
            (
                key,
                city,
                address,
                property_type,
                complex_name,
                json.dumps(payload, ensure_ascii=False),
                len(payload),
                now,
                now,
            ),
        )
        conn.commit()
    finally:
        conn.close()


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")
