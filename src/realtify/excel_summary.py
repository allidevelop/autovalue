from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from realtify.excel_tools import ExcelApp, excel_path


def read_excel_report_values(path: Path | None) -> dict[str, Any]:
    if not path:
        return {}
    workbook_path = path.resolve()
    if not workbook_path.exists():
        return {}

    with ExcelApp(visible=False) as excel:
        wb = excel.Workbooks.Open(excel_path(workbook_path), 0, True)
        try:
            ws = _calculation_sheet(wb)
            target_area = _cell_value(ws, 3, 7)
            average_usd_m2 = _cell_value(ws, 37, 3)
            median_usd_m2 = _cell_value(ws, 38, 3)
            market_value_raw = _cell_value(ws, 44, 6)
            nbu_rate = _rate_from_formula(ws.Cells(44, 6).Formula)
            if market_value_raw is None and median_usd_m2 is not None and target_area is not None and nbu_rate is not None:
                market_value_raw = float(median_usd_m2) * float(target_area) * float(nbu_rate)
            median_uah_m2 = None
            if median_usd_m2 is not None and nbu_rate is not None:
                median_uah_m2 = float(median_usd_m2) * float(nbu_rate)
            rounded = _round_money(market_value_raw)
            return {
                "average_price_usd_m2": average_usd_m2,
                "median_price_usd_m2": median_usd_m2,
                "nbu_rate": nbu_rate,
                "median_price_uah_m2": median_uah_m2,
                "market_value_uah": market_value_raw,
                "market_value_uah_rounded": rounded,
                "market_value_uah_words": number_to_ukrainian_words(int(rounded)) if rounded is not None else None,
            }
        finally:
            wb.Close(False)


def _calculation_sheet(wb: Any) -> Any:
    for ws in wb.Worksheets:
        name = str(ws.Name).casefold()
        if "\u0440\u043e\u0437\u0440\u0430\u0445\u0443\u043d\u043e\u043a" in name and "\u0432\u0430\u0440\u0442" in name:
            return ws
    return wb.Worksheets(3)


def _cell_value(ws: Any, row: int, col: int) -> float | None:
    value = ws.Cells(row, col).Value
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _rate_from_formula(formula: Any) -> float | None:
    if not formula:
        return None
    numbers = [float(value.replace(",", ".")) for value in re.findall(r"(?<![A-Z])(\d+[.,]\d+)", str(formula))]
    return numbers[-1] if numbers else None


def _round_money(value: float | None) -> float | None:
    if value is None:
        return None
    return round(float(value) / 100) * 100


def number_to_ukrainian_words(value: int) -> str:
    if value == 0:
        return "\u043d\u0443\u043b\u044c"
    groups = [
        ("", "", "", False),
        ("\u0442\u0438\u0441\u044f\u0447\u0430", "\u0442\u0438\u0441\u044f\u0447\u0456", "\u0442\u0438\u0441\u044f\u0447", True),
        ("\u043c\u0456\u043b\u044c\u0439\u043e\u043d", "\u043c\u0456\u043b\u044c\u0439\u043e\u043d\u0438", "\u043c\u0456\u043b\u044c\u0439\u043e\u043d\u0456\u0432", False),
        ("\u043c\u0456\u043b\u044c\u044f\u0440\u0434", "\u043c\u0456\u043b\u044c\u044f\u0440\u0434\u0438", "\u043c\u0456\u043b\u044c\u044f\u0440\u0434\u0456\u0432", False),
    ]
    parts: list[str] = []
    group_index = 0
    while value > 0:
        chunk = value % 1000
        if chunk:
            noun_forms = groups[group_index]
            words = _chunk_to_words(chunk, feminine=noun_forms[3])
            noun = _choose_plural(chunk, noun_forms[0], noun_forms[1], noun_forms[2])
            if noun:
                words.append(noun)
            parts.insert(0, " ".join(words))
        value //= 1000
        group_index += 1
    return " ".join(parts)


def _chunk_to_words(value: int, *, feminine: bool) -> list[str]:
    hundreds = [
        "",
        "\u0441\u0442\u043e",
        "\u0434\u0432\u0456\u0441\u0442\u0456",
        "\u0442\u0440\u0438\u0441\u0442\u0430",
        "\u0447\u043e\u0442\u0438\u0440\u0438\u0441\u0442\u0430",
        "\u043f\u2019\u044f\u0442\u0441\u043e\u0442",
        "\u0448\u0456\u0441\u0442\u0441\u043e\u0442",
        "\u0441\u0456\u043c\u0441\u043e\u0442",
        "\u0432\u0456\u0441\u0456\u043c\u0441\u043e\u0442",
        "\u0434\u0435\u0432\u2019\u044f\u0442\u0441\u043e\u0442",
    ]
    tens = [
        "",
        "",
        "\u0434\u0432\u0430\u0434\u0446\u044f\u0442\u044c",
        "\u0442\u0440\u0438\u0434\u0446\u044f\u0442\u044c",
        "\u0441\u043e\u0440\u043e\u043a",
        "\u043f\u2019\u044f\u0442\u0434\u0435\u0441\u044f\u0442",
        "\u0448\u0456\u0441\u0442\u0434\u0435\u0441\u044f\u0442",
        "\u0441\u0456\u043c\u0434\u0435\u0441\u044f\u0442",
        "\u0432\u0456\u0441\u0456\u043c\u0434\u0435\u0441\u044f\u0442",
        "\u0434\u0435\u0432\u2019\u044f\u043d\u043e\u0441\u0442\u043e",
    ]
    teens = [
        "\u0434\u0435\u0441\u044f\u0442\u044c",
        "\u043e\u0434\u0438\u043d\u0430\u0434\u0446\u044f\u0442\u044c",
        "\u0434\u0432\u0430\u043d\u0430\u0434\u0446\u044f\u0442\u044c",
        "\u0442\u0440\u0438\u043d\u0430\u0434\u0446\u044f\u0442\u044c",
        "\u0447\u043e\u0442\u0438\u0440\u043d\u0430\u0434\u0446\u044f\u0442\u044c",
        "\u043f\u2019\u044f\u0442\u043d\u0430\u0434\u0446\u044f\u0442\u044c",
        "\u0448\u0456\u0441\u0442\u043d\u0430\u0434\u0446\u044f\u0442\u044c",
        "\u0441\u0456\u043c\u043d\u0430\u0434\u0446\u044f\u0442\u044c",
        "\u0432\u0456\u0441\u0456\u043c\u043d\u0430\u0434\u0446\u044f\u0442\u044c",
        "\u0434\u0435\u0432\u2019\u044f\u0442\u043d\u0430\u0434\u0446\u044f\u0442\u044c",
    ]
    ones_m = ["", "\u043e\u0434\u0438\u043d", "\u0434\u0432\u0430", "\u0442\u0440\u0438", "\u0447\u043e\u0442\u0438\u0440\u0438", "\u043f\u2019\u044f\u0442\u044c", "\u0448\u0456\u0441\u0442\u044c", "\u0441\u0456\u043c", "\u0432\u0456\u0441\u0456\u043c", "\u0434\u0435\u0432\u2019\u044f\u0442\u044c"]
    ones_f = ["", "\u043e\u0434\u043d\u0430", "\u0434\u0432\u0456", "\u0442\u0440\u0438", "\u0447\u043e\u0442\u0438\u0440\u0438", "\u043f\u2019\u044f\u0442\u044c", "\u0448\u0456\u0441\u0442\u044c", "\u0441\u0456\u043c", "\u0432\u0456\u0441\u0456\u043c", "\u0434\u0435\u0432\u2019\u044f\u0442\u044c"]

    words: list[str] = []
    h = value // 100
    if h:
        words.append(hundreds[h])
    remainder = value % 100
    if 10 <= remainder <= 19:
        words.append(teens[remainder - 10])
        return words
    t = remainder // 10
    if t:
        words.append(tens[t])
    o = remainder % 10
    if o:
        words.append((ones_f if feminine else ones_m)[o])
    return words


def _choose_plural(value: int, one: str, few: str, many: str) -> str:
    if not one:
        return ""
    last_two = value % 100
    last = value % 10
    if 11 <= last_two <= 14:
        return many
    if last == 1:
        return one
    if 2 <= last <= 4:
        return few
    return many
