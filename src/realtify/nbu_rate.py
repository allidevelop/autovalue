"""Офіційний курс НБУ на дату оцінки.

Детермінований, без нейромереж: публічний API НБУ (без ключа) →
курс гривні за 1 одиницю валюти на конкретну дату. Результат кешується
на диску, щоб та сама дата не смикала API повторно.

Логіка клієнта: «дата оцінки — з реєстру, і відповідно курс НБУ на цю дату».
"""

from __future__ import annotations

import json
import os
import sys
from datetime import date, timedelta
from pathlib import Path
from typing import Any

import requests

from realtify.paths import PROJECT_ROOT

NBU_EXCHANGE_URL = "https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange"
_CACHE_PATH = PROJECT_ROOT / "data" / "nbu_cache" / "rates.json"
_TIMEOUT_SECONDS = 15
_FETCH_ATTEMPTS = 3  # ретрай на транзиентний мережевий збій (холодний DNS/TLS)
# На вихідні/свята НБУ не публікує нового курсу — крок назад до останнього опублікованого.
_MAX_LOOKBACK_DAYS = 6


class _NetworkError(RuntimeError):
    """Збій мережі/HTTP — на відміну від «на цю дату курсу немає»."""


def usd_uah_rate(target_day: date, *, valcode: str = "USD") -> float | None:
    """Курс НБУ (грн за 1 одиницю `valcode`) на `target_day`.

    Якщо на дату курсу немає (вихідний/свято) — береться останній опублікований
    курс (крок назад, макс. 6 днів). Повертає None, якщо мережа недоступна
    і в кеші запису немає — викликач має відкотитись на дефолтний курс.
    """
    cache = _load_cache()
    orig_key = f"{valcode}:{target_day.isoformat()}"
    cached = cache.get(orig_key)
    if cached is not None:
        try:
            return float(cached)
        except (TypeError, ValueError):
            return None

    day = target_day
    try:
        for _ in range(_MAX_LOOKBACK_DAYS + 1):
            rate = _fetch(valcode, day)
            if rate is not None:
                cache[orig_key] = rate
                _save_cache(cache)
                return rate
            # Порожня відповідь = на цю дату курсу немає → крок на день назад.
            day = day - timedelta(days=1)
    except _NetworkError:
        # Мережа недоступна — не кешуємо «міс», щоб повторити пізніше.
        return None
    return None


def _fetch(valcode: str, day: date) -> float | None:
    # `&json` — прапорець без значення (саме так очікує API НБУ).
    url = f"{NBU_EXCHANGE_URL}?valcode={valcode}&date={day.strftime('%Y%m%d')}&json"
    last_exc: Exception | None = None
    payload: Any = None
    for _ in range(_FETCH_ATTEMPTS):
        try:
            response = requests.get(
                url,
                timeout=_TIMEOUT_SECONDS,
                headers={"User-Agent": "realtify-autovalue/1.0"},
            )
            response.raise_for_status()
            payload = response.json()
            break
        except Exception as exc:  # noqa: BLE001 — транзиентний збій → ретрай
            last_exc = exc
    else:
        # Діагностика: фіксуємо причину збою (інакше None невідрізнюваний від «немає курсу»).
        print(f"[nbu_rate] network failure for {day.isoformat()}: {last_exc}", file=sys.stderr)
        raise _NetworkError(str(last_exc)) from last_exc
    if isinstance(payload, list) and payload:
        rate = payload[0].get("rate") if isinstance(payload[0], dict) else None
        if rate is not None:
            try:
                value = float(rate)
            except (TypeError, ValueError):
                return None
            # Відкидаємо нереальні значення (0/від'ємне) — інакше market_value=0 з official=True.
            if value <= 0:
                return None
            return value
    return None


def _load_cache() -> dict[str, Any]:
    try:
        with _CACHE_PATH.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
        return data if isinstance(data, dict) else {}
    except Exception:  # noqa: BLE001 — немає кешу/битий кеш → починаємо з порожнього
        return {}


def _save_cache(cache: dict[str, Any]) -> None:
    try:
        _CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
        # Атомарний запис: temp у тій самій теці → os.replace, щоб краш не лишив битий JSON.
        tmp = _CACHE_PATH.with_name(f"{_CACHE_PATH.name}.tmp")
        with tmp.open("w", encoding="utf-8") as handle:
            json.dump(cache, handle, ensure_ascii=False, indent=0, sort_keys=True)
        os.replace(tmp, _CACHE_PATH)
    except Exception:  # noqa: BLE001 — кеш не критичний для основного потоку
        pass
