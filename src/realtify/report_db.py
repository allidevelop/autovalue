"""База аналогів (порівнянь) із готових звітів оцінки + знайдених при пошуку.

Нормалізована таблиця `report_comparables` у тому ж `data/analog_cache/cache.db`:
один рядок = один унікальний аналог (оголошення-пропозиція). Дедуплікація між
звітами за `dedup_key`. Пошук для нової оцінки: точний будинок → той самий ЖК →
місто (розширення). Поля повторюють таблицю-скриншот звіту (адреса, площа, поверх,
ціна USD, $/м², локація, клас, оздоблення, термін, джерело).

Власник бази (CRUD) — веб-сторінка realtifysaas через API autovalue.
"""

from __future__ import annotations

import sqlite3
from datetime import date, datetime
from pathlib import Path
from typing import Any

from realtify.analog_cache import DB_PATH, address_key
from realtify.models import Comparable

# Колонки, які приходять/повертаються (без службових id/created/updated).
FIELDS: tuple[str, ...] = (
    "address_key", "city", "address", "complex_name", "property_type",
    "area_m2", "price_usd", "price_per_m2_usd", "floor_or_level", "rooms",
    "location_quality", "building_class", "condition", "delivery_date",
    "listing_date", "source_key", "report_id", "source_url",
)


def _connect() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS report_comparables (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            dedup_key        TEXT UNIQUE,
            address_key      TEXT,
            city             TEXT,
            address          TEXT,
            complex_name     TEXT,
            property_type    TEXT,
            area_m2          REAL,
            price_usd        REAL,
            price_per_m2_usd REAL,
            floor_or_level   TEXT,
            rooms            INTEGER,
            location_quality TEXT,
            building_class   TEXT,
            condition        TEXT,
            delivery_date    TEXT,
            listing_date     TEXT,
            source_key       TEXT,
            report_id        TEXT,
            source_url       TEXT,
            created_at       TEXT,
            updated_at       TEXT
        )
        """
    )
    for col in ("address_key", "complex_name", "city"):
        conn.execute(f"CREATE INDEX IF NOT EXISTS ix_rc_{col} ON report_comparables({col})")
    return conn


# ── запис/імпорт ─────────────────────────────────────────────────────────────

def compute_dedup_key(address_key_value: str | None, area_m2: Any, price_usd: Any, floor: Any) -> str:
    a = _round(area_m2, 1)
    p = _round(price_usd, 0)
    f = str(floor or "").strip().lower()
    return f"{address_key_value or ''}|{a}|{p}|{f}"


def upsert_many(rows: list[dict[str, Any]]) -> dict[str, int]:
    """Idempotent upsert за dedup_key. Повертає {inserted, updated}."""
    now = datetime.now().isoformat(timespec="seconds")
    inserted = updated = 0
    conn = _connect()
    try:
        for raw in rows:
            row = _normalize_row(raw)
            dedup = compute_dedup_key(row["address_key"], row["area_m2"], row["price_usd"], row["floor_or_level"])
            existing = conn.execute("SELECT id FROM report_comparables WHERE dedup_key = ?", (dedup,)).fetchone()
            cols = list(FIELDS)
            values = [row.get(c) for c in cols]
            if existing:
                assignments = ", ".join(f"{c} = ?" for c in cols)
                conn.execute(
                    f"UPDATE report_comparables SET {assignments}, updated_at = ? WHERE dedup_key = ?",
                    (*values, now, dedup),
                )
                updated += 1
            else:
                placeholders = ", ".join(["?"] * (len(cols) + 3))
                conn.execute(
                    f"INSERT INTO report_comparables (dedup_key, {', '.join(cols)}, created_at, updated_at) "
                    f"VALUES ({placeholders})",
                    (dedup, *values, now, now),
                )
                inserted += 1
        conn.commit()
    finally:
        conn.close()
    return {"inserted": inserted, "updated": updated}


# ── пошук для оцінки ─────────────────────────────────────────────────────────

def find_comparables(
    target: dict[str, Any],
    *,
    as_of_date: date | None = None,
    min_count: int = 5,
    property_type: str | None = None,
) -> tuple[list[Comparable], str]:
    """Аналоги з бази для об'єкта: будинок → ЖК → місто. Повертає (список, рівень)."""
    ptype = str(property_type or target.get("property_type") or "apartment")
    key = address_key(
        city=target.get("city"),
        address=target.get("address"),
        property_type=ptype,
        complex_name=target.get("complex_name"),
    )
    complex_name = _norm(target.get("complex_name"))
    city = _norm(target.get("city"))
    target_area = _to_float(target.get("area_m2"))

    conn = _connect()
    try:
        seen: set[int] = set()
        tier = "none"
        rows: list[sqlite3.Row] = []

        building = conn.execute(
            "SELECT * FROM report_comparables WHERE address_key = ?", (key,)
        ).fetchall()
        if building:
            tier = "building"
            rows.extend(building)

        # SQLite lower()/COLLATE NOCASE — лише ASCII; кирилицю нормалізуємо в Python.
        if len(rows) < min_count and complex_name:
            pool = conn.execute(
                "SELECT * FROM report_comparables WHERE property_type = ? AND complex_name IS NOT NULL",
                (ptype,),
            ).fetchall()
            have = {x["id"] for x in rows}
            new = [r for r in pool if _norm(r["complex_name"]) == complex_name and r["id"] not in have]
            if new:
                if not building:
                    tier = "complex"
                rows.extend(new)

        if len(rows) < min_count and city:
            pool = conn.execute(
                "SELECT * FROM report_comparables WHERE property_type = ? AND city IS NOT NULL",
                (ptype,),
            ).fetchall()
            have = {x["id"] for x in rows}
            new = [r for r in pool if _norm(r["city"]) == city and r["id"] not in have]
            if new:
                if not rows:
                    tier = "city"
                rows.extend(new)
    finally:
        conn.close()

    ranked = sorted(rows, key=lambda r: _rank(r, as_of_date, target_area))
    out: list[Comparable] = []
    for r in ranked:
        if r["id"] in seen:
            continue
        seen.add(r["id"])
        out.append(_row_to_comparable(r))
    return out, tier


# ── CRUD (для веб-сторінки) ──────────────────────────────────────────────────

def list_groups() -> list[dict[str, Any]]:
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT address_key, "
            "MAX(address) AS address, MAX(complex_name) AS complex_name, MAX(city) AS city, "
            "COUNT(*) AS count, MAX(updated_at) AS updated_at "
            "FROM report_comparables GROUP BY address_key "
            "ORDER BY MAX(complex_name) IS NULL, MAX(complex_name), MAX(address)"
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def list_items(address_key_value: str) -> list[dict[str, Any]]:
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM report_comparables WHERE address_key = ? ORDER BY area_m2", (address_key_value,)
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get(item_id: int) -> dict[str, Any] | None:
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM report_comparables WHERE id = ?", (item_id,)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def create(payload: dict[str, Any]) -> dict[str, Any]:
    row = _normalize_row(payload, recompute_key=True)
    now = datetime.now().isoformat(timespec="seconds")
    dedup = compute_dedup_key(row["address_key"], row["area_m2"], row["price_usd"], row["floor_or_level"])
    cols = list(FIELDS)
    conn = _connect()
    try:
        placeholders = ", ".join(["?"] * (len(cols) + 3))
        cur = conn.execute(
            f"INSERT INTO report_comparables (dedup_key, {', '.join(cols)}, created_at, updated_at) "
            f"VALUES ({placeholders})",
            (dedup, *[row.get(c) for c in cols], now, now),
        )
        conn.commit()
        new_id = cur.lastrowid
    finally:
        conn.close()
    return get(int(new_id)) or {}


def update(item_id: int, payload: dict[str, Any]) -> dict[str, Any] | None:
    current = get(item_id)
    if not current:
        return None
    merged = {**current, **{k: v for k, v in payload.items() if k in FIELDS}}
    row = _normalize_row(merged, recompute_key=True)
    now = datetime.now().isoformat(timespec="seconds")
    dedup = compute_dedup_key(row["address_key"], row["area_m2"], row["price_usd"], row["floor_or_level"])
    cols = list(FIELDS)
    assignments = ", ".join(f"{c} = ?" for c in cols)
    conn = _connect()
    try:
        conn.execute(
            f"UPDATE report_comparables SET dedup_key = ?, {assignments}, updated_at = ? WHERE id = ?",
            (dedup, *[row.get(c) for c in cols], now, item_id),
        )
        conn.commit()
    finally:
        conn.close()
    return get(item_id)


def delete(item_id: int) -> bool:
    conn = _connect()
    try:
        cur = conn.execute("DELETE FROM report_comparables WHERE id = ?", (item_id,))
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


# ── допоміжне ────────────────────────────────────────────────────────────────

def _normalize_row(raw: dict[str, Any], *, recompute_key: bool = False) -> dict[str, Any]:
    row = {c: raw.get(c) for c in FIELDS}
    row["area_m2"] = _to_float(row.get("area_m2"))
    row["price_usd"] = _to_float(row.get("price_usd"))
    row["price_per_m2_usd"] = _to_float(row.get("price_per_m2_usd"))
    # $/м² завжди узгоджений з ціною/площею (інакше при правці ціни лишається старе).
    if row["price_usd"] and row["area_m2"]:
        row["price_per_m2_usd"] = round(row["price_usd"] / row["area_m2"], 2)
    row["rooms"] = _to_int(row.get("rooms"))
    row["property_type"] = str(row.get("property_type") or "apartment")
    if recompute_key or not row.get("address_key"):
        row["address_key"] = address_key(
            city=row.get("city"), address=row.get("address"),
            property_type=row["property_type"], complex_name=row.get("complex_name"),
        )
    if not row.get("source_url"):
        row["source_url"] = "https://report.local/manual"
    if not row.get("source_key"):
        row["source_key"] = "manual"
    ld = row.get("listing_date")
    if isinstance(ld, (date, datetime)):
        row["listing_date"] = ld.isoformat()[:10]
    return row


def _row_to_comparable(row: sqlite3.Row) -> Comparable:
    collected = _parse_dt(row["listing_date"]) or datetime.now()
    return Comparable.model_validate({
        "source_url": row["source_url"] or "https://report.local/x",
        "source_key": row["source_key"],
        "source_name": "Архів оцінок" if row["source_key"] == "report_archive" else None,
        "property_type": row["property_type"] or "apartment",
        "transaction_type": "sale",
        "address": row["address"],
        "city": row["city"],
        "complex_name": row["complex_name"],
        "area_m2": row["area_m2"],
        "price_usd": row["price_usd"],
        "price_per_m2_usd": row["price_per_m2_usd"],
        "floor_or_level": row["floor_or_level"],
        "rooms": row["rooms"],
        "location_quality": row["location_quality"],
        "building_class": row["building_class"],
        "condition": row["condition"],
        "delivery_date": row["delivery_date"],
        "collected_at": collected,
    })


def _rank(row: sqlite3.Row, as_of: date | None, target_area: float | None) -> tuple[float, float]:
    date_delta = 0.0
    ld = _parse_dt(row["listing_date"])
    if as_of and ld:
        date_delta = abs((ld.date() - as_of).days)
    area_delta = 0.0
    if target_area and row["area_m2"]:
        area_delta = abs(row["area_m2"] - target_area)
    return (date_delta, area_delta)


def _round(value: Any, ndigits: int) -> Any:
    f = _to_float(value)
    return round(f, ndigits) if f is not None else ""


def _norm(value: Any) -> str:
    return str(value).strip().lower() if value not in (None, "") else ""


def _to_float(value: Any) -> float | None:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _to_int(value: Any) -> int | None:
    f = _to_float(value)
    return int(f) if f is not None else None


def _parse_dt(value: Any) -> datetime | None:
    if not value:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(str(value)[:10], fmt)
        except ValueError:
            continue
    return None
