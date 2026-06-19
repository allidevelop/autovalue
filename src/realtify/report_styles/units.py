"""Конвертації одиниць для звітного рендеру (спільні для HTML/PDF та .docx).

EMU — English Metric Units (OOXML / python-docx). Базові співвідношення:
1 inch = 914400 EMU = 72 pt = 25.4 mm; 1 mm = 36000 EMU; 1 pt = 12700 EMU.

Геометрію (ширини колонок, поля) задаємо в МІЛІМЕТРАХ один раз у style_spec —
обидва рендерери конвертують її однаково, тож таблиці й поля збігаються.
"""
from __future__ import annotations

EMU_PER_INCH = 914400
EMU_PER_MM = 36000
EMU_PER_PT = 12700
PT_PER_INCH = 72.0
MM_PER_INCH = 25.4


def mm_to_emu(mm: float) -> int:
    return round(mm * EMU_PER_MM)


def pt_to_emu(pt: float) -> int:
    return round(pt * EMU_PER_PT)


def inch_to_emu(inch: float) -> int:
    return round(inch * EMU_PER_INCH)


def mm_to_pt(mm: float) -> float:
    return mm / MM_PER_INCH * PT_PER_INCH


def pt_to_mm(pt: float) -> float:
    return pt / PT_PER_INCH * MM_PER_INCH


def emu_to_mm(emu: float) -> float:
    return emu / EMU_PER_MM


def emu_to_inch(emu: float) -> float:
    return emu / EMU_PER_INCH
