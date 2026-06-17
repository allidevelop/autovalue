from __future__ import annotations

import json
import math
import statistics
from datetime import date, datetime
from pathlib import Path
from typing import Any

from realtify.models import Comparable


SIDECAR_SCHEMA_VERSION = 1
DEFAULT_NBU_RATE = 37.6166
DEFAULT_T_CONFIDENCE = 2.776


def sidecar_path(excel_path: Path) -> Path:
    return excel_path.with_suffix(excel_path.suffix + ".realtify.json")


def load_excel_sidecar(excel_path: Path | None) -> dict[str, Any] | None:
    if not excel_path:
        return None
    path = sidecar_path(excel_path)
    if not path.exists():
        return None
    with path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)
    if payload.get("schema_version") != SIDECAR_SCHEMA_VERSION:
        return None
    return payload


def write_excel_sidecar(excel_path: Path, payload: dict[str, Any]) -> Path:
    path = sidecar_path(excel_path)
    data = {
        "schema_version": SIDECAR_SCHEMA_VERSION,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        **payload,
    }
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2, default=_json_default), encoding="utf-8")
    return path


def _json_default(value: Any) -> str:
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return str(value)


def sidecar_summary_values(excel_path: Path | None) -> dict[str, Any]:
    payload = load_excel_sidecar(excel_path)
    if not payload:
        return {}
    summary = payload.get("summary")
    return summary if isinstance(summary, dict) else {}


def sidecar_adjustment_rows(excel_path: Path | None) -> dict[int, list[str]]:
    payload = load_excel_sidecar(excel_path)
    if not payload:
        return {}
    raw_rows = payload.get("adjustment_rows")
    if not isinstance(raw_rows, dict):
        return {}
    rows: dict[int, list[str]] = {}
    for key, value in raw_rows.items():
        if isinstance(value, list):
            rows[int(key)] = [str(item) for item in value]
    return rows


def sidecar_visible_text(excel_path: Path | None) -> str:
    payload = load_excel_sidecar(excel_path)
    if not payload:
        return ""
    value = payload.get("visible_text")
    return str(value or "")


def build_calculation_sidecar_payload(
    *,
    excel_path: Path,
    engine: str,
    profile_name: str,
    candidates: list[Comparable],
    target: dict[str, Any] | None,
    template_rows: dict[int, list[Any]] | None = None,
    nbu_rate: float | None = DEFAULT_NBU_RATE,
) -> dict[str, Any]:
    if nbu_rate is None:
        nbu_rate = DEFAULT_NBU_RATE
    target = target or {}
    target_area = _to_float(target.get("area_m2"))
    selected = candidates[:5]
    candidate_rows = [_candidate_inputs(candidate) for candidate in selected]
    while len(candidate_rows) < 5:
        candidate_rows.append({})

    rows: dict[int, list[Any]] = {}
    for row_index in range(15, 44):
        rows[row_index] = _base_row(template_rows, row_index)

    price_m2 = [_price_per_m2(item) for item in candidate_rows]
    areas = [_to_float(item.get("area_m2")) for item in candidate_rows]

    _set_row_values(rows, 16, 4, price_m2)
    _set_row_values(rows, 17, 3, [target.get("address")] + [item.get("address") for item in candidate_rows])
    _set_row_values(rows, 18, 4, [_row_value(rows, 18, col) for col in range(4, 9)])
    _set_row_values(rows, 19, 4, [_numeric_row_value(rows, 19, col, 0.0) for col in range(4, 9)])

    row20 = [_pct_adjust(value, _numeric_row_value(rows, 19, col, 0.0)) for col, value in enumerate(price_m2, start=4)]
    _set_row_values(rows, 20, 4, row20)
    _set_row_values(rows, 21, 4, [_numeric_row_value(rows, 21, col, 0.0) for col in range(4, 9)])

    row22 = [_pct_adjust(value, _numeric_row_value(rows, 21, col, 0.0)) for col, value in enumerate(row20, start=4)]
    _set_row_values(rows, 22, 4, row22)
    _set_row_values(rows, 23, 4, [_numeric_row_value(rows, 23, col, 0.0) for col in range(4, 9)])

    row24 = [_pct_adjust(value, _numeric_row_value(rows, 23, col, 0.0)) for col, value in enumerate(row22, start=4)]
    _set_row_values(rows, 24, 4, row24)
    _set_row_values(rows, 25, 4, [_numeric_row_value(rows, 25, col, 0.0) for col in range(4, 9)])

    row26 = [_pct_adjust(value, _numeric_row_value(rows, 25, col, 0.0)) for col, value in enumerate(row24, start=4)]
    _set_row_values(rows, 26, 4, row26)

    area_coefficients = [_area_coefficient(target_area, area) for area in areas]
    _set_row_values(rows, 27, 4, area_coefficients)

    row28 = [_multiply(value, coeff) for value, coeff in zip(row26, area_coefficients)]
    _set_row_values(rows, 28, 4, row28)
    _set_row_values(rows, 29, 3, [target.get("wall_material") or _row_value(rows, 29, 3)] + [_row_value(rows, 29, col) for col in range(4, 9)])
    _set_row_values(rows, 30, 3, [target.get("building_class") or _row_value(rows, 30, 3)] + [_numeric_row_value(rows, 30, col, 0.0) for col in range(4, 9)])

    row31 = [_pct_adjust(value, _numeric_row_value(rows, 30, col, 0.0)) for col, value in enumerate(row28, start=4)]
    _set_row_values(rows, 31, 4, row31)
    _set_row_values(rows, 32, 4, [_numeric_row_value(rows, 32, col, 0.0) for col in range(4, 9)])

    row33 = [_pct_adjust(value, _numeric_row_value(rows, 32, col, 0.0)) for col, value in enumerate(row31, start=4)]
    _set_row_values(rows, 33, 4, row33)
    _set_row_values(rows, 34, 4, [_numeric_row_value(rows, 34, col, 0.0) for col in range(4, 9)])

    row35 = [_multiply(value, _numeric_row_value(rows, 34, col, 0.0) / 100.0) for col, value in enumerate(row33, start=4)]
    _set_row_values(rows, 35, 4, row35)

    row36 = [_sum_optional(value, extra) for value, extra in zip(row33, row35)]
    _set_row_values(rows, 36, 4, row36)

    average = _mean([value for value in row36[:3] if value is not None])
    median = _median([value for value in row36 if value is not None])
    stdev = _stdev([value for value in row36[:3] if value is not None])
    variation = stdev / average if average else None
    lower = average - (DEFAULT_T_CONFIDENCE * stdev / 2.0) if average is not None and stdev is not None else None
    upper = average + (DEFAULT_T_CONFIDENCE * stdev / 2.0) if average is not None and stdev is not None else None
    interval = (upper - lower) / 2.0 if upper is not None and lower is not None else None

    _set_row_values(rows, 37, 3, [average])
    _set_row_values(rows, 38, 3, [median])
    _set_row_values(rows, 39, 3, [stdev])
    _set_row_values(rows, 40, 3, [variation])
    _set_row_values(rows, 41, 3, [interval])
    _set_row_values(rows, 42, 3, [lower])
    _set_row_values(rows, 43, 3, [upper])

    market_value_usd = _multiply(median, target_area)
    market_value_uah = _multiply(market_value_usd, nbu_rate)
    market_value_rounded = _round_money(market_value_uah)

    formatted_rows = {
        str(row_index): _format_adjustment_row(row_index, rows[row_index])
        for row_index in range(15, 44)
    }
    raw_rows = {
        str(row_index): rows[row_index]
        for row_index in range(15, 44)
    }
    visible_values: list[str] = []
    for value in (
        target.get("address"),
        target.get("area_m2"),
        target.get("floor_or_level"),
        target.get("complex_name"),
    ):
        if value not in (None, ""):
            visible_values.append(str(value))
    for item in candidate_rows:
        for value in (item.get("address"), item.get("area_m2"), item.get("price_usd")):
            if value not in (None, ""):
                visible_values.append(str(value))
    for row in formatted_rows.values():
        visible_values.extend(str(value) for value in row if str(value).strip())

    return {
        "excel_path": str(excel_path),
        "engine": engine,
        "profile": profile_name,
        "summary": {
            "average_price_usd_m2": average,
            "median_price_usd_m2": median,
            "nbu_rate": nbu_rate,
            "median_price_uah_m2": _multiply(median, nbu_rate),
            "market_value_usd": market_value_usd,
            "market_value_uah": market_value_uah,
            "market_value_uah_rounded": market_value_rounded,
        },
        "adjustment_rows": formatted_rows,
        "raw_adjustment_rows": raw_rows,
        "visible_text": "\n".join(visible_values),
    }


def _candidate_inputs(candidate: Comparable) -> dict[str, Any]:
    address = candidate.address
    if address and candidate.complex_name and candidate.complex_name not in address:
        address = f"{address}, {candidate.complex_name}"
    return {
        "address": address,
        "area_m2": candidate.area_m2,
        "price_usd": candidate.price_usd if candidate.price_usd is not None else candidate.price,
        "price_per_m2_usd": candidate.price_per_m2_usd,
    }


def _base_row(template_rows: dict[int, list[Any]] | None, row_index: int) -> list[Any]:
    if template_rows and row_index in template_rows:
        row = list(template_rows[row_index][:8])
    else:
        row = [""] * 8
    while len(row) < 8:
        row.append("")
    return row


def _set_row_values(rows: dict[int, list[Any]], row_index: int, start_col: int, values: list[Any]) -> None:
    row = rows.setdefault(row_index, [""] * 8)
    while len(row) < 8:
        row.append("")
    for offset, value in enumerate(values):
        col_index = start_col + offset
        if 1 <= col_index <= len(row):
            row[col_index - 1] = value


def _row_value(rows: dict[int, list[Any]], row_index: int, col_index: int) -> Any:
    try:
        return rows[row_index][col_index - 1]
    except (KeyError, IndexError):
        return None


def _numeric_row_value(rows: dict[int, list[Any]], row_index: int, col_index: int, default: float) -> float:
    value = _to_float(_row_value(rows, row_index, col_index))
    return default if value is None else value


def _price_per_m2(candidate: dict[str, Any]) -> float | None:
    value = _to_float(candidate.get("price_per_m2_usd"))
    if value is not None:
        return value
    price = _to_float(candidate.get("price_usd"))
    area = _to_float(candidate.get("area_m2"))
    if price is None or not area:
        return None
    return price / area


def _pct_adjust(value: float | None, pct: float) -> float | None:
    if value is None:
        return None
    return value * ((100.0 + pct) / 100.0)


def _multiply(left: float | None, right: float | None) -> float | None:
    if left is None or right is None:
        return None
    return left * right


def _sum_optional(left: float | None, right: float | None) -> float | None:
    if left is None and right is None:
        return None
    return (left or 0.0) + (right or 0.0)


def _area_coefficient(target_area: float | None, candidate_area: float | None) -> float | None:
    if not target_area or not candidate_area:
        return None
    return (candidate_area / target_area) ** (0.1 * target_area / (candidate_area + target_area))


def _mean(values: list[float]) -> float | None:
    return statistics.fmean(values) if values else None


def _median(values: list[float]) -> float | None:
    return statistics.median(values) if values else None


def _stdev(values: list[float]) -> float | None:
    return statistics.stdev(values) if len(values) >= 2 else None


def _round_money(value: float | None) -> float | None:
    if value is None:
        return None
    return round(float(value) / 100) * 100


def _format_adjustment_row(row_index: int, row: list[Any]) -> list[str]:
    formatted: list[str] = []
    for col_index, value in enumerate(row[:8], start=1):
        formatted.append(_format_adjustment_cell(row_index, col_index, value))
    return formatted


def _format_adjustment_cell(row_index: int, col_index: int, value: Any) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    number = _to_float(value)
    if number is None:
        return str(value)
    if row_index in {27, 40}:
        return f"{number:.2f}"
    if row_index in {19, 21, 23, 25, 30, 32, 34}:
        return _format_trimmed_number(number)
    return _format_money_number(number)


def _format_money_number(value: float) -> str:
    return f"{value:,.0f}".replace(",", " ")


def _format_trimmed_number(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.2f}".rstrip("0").rstrip(".")


def _to_float(value: Any) -> float | None:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("\xa0", " ").replace(" ", "").replace(",", ".")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None
