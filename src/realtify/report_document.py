"""Збірка структурованого документа звіту (schema JSON) з даних об'єкта.

ПЕРЕВИКОРИСТОВУЄ існуючу збірку даних — не пересчитує: значення беруться з
report_generator.build_report_values(), розрахункова таблиця — з Excel-сайдкара
(excel_sidecar.sidecar_adjustment_rows). Кожне динамічне значення стає
variableField {field, source}; таблиця поправок — locked-нода з сайдкара.

ФАЗА 1 (вертикальний зріз): розділ 5.1 Місцезнаходження + Таблиця 7.3 поправок.
Решта блоків додаються у Фазі 2.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from realtify import report_schema as S
from realtify import report_styles as styles
from realtify.excel_sidecar import sidecar_adjustment_rows
from realtify.report_generator import (
    _load_candidates, _load_intake, _load_yaml, _resolve_optional, build_report_values,
)
from realtify.excel_summary import read_excel_report_values

# Порядок 27 видимих рядків таблиці поправок (= report_generator.word_to_excel_rows).
_ADJ_ROW_ORDER = [15, 16, 17, 18, 19, 20, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32,
                  33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43]
_ADJ_HEADER = ["Показник", "Одиниці", "Об'єкт оцінки", "Аналог №1", "Аналог №2",
               "Аналог №3", "Аналог №4", "Аналог №5"]


def _adjustment_table_node(excel_path: Path | None, spec: dict[str, Any]) -> dict:
    rows_map = sidecar_adjustment_rows(excel_path) if excel_path else {}
    rows: list[list[str]] = []
    for idx in _ADJ_ROW_ORDER:
        r = rows_map.get(idx) or rows_map.get(str(idx))
        if r:
            rows.append([str(c) for c in r[:8]])
    cols = styles.table_style(spec, "adjustment_73")["columnsMm"]
    return S.table("adjustment_73", _ADJ_HEADER, rows, cols)


def build_report_document(*, intake, task: dict, candidates: list, excel_path: Path | None) -> dict:
    """Будує повний schema-JSON документа звіту з даних об'єкта (Фаза 2):
    імпортує структуру вже-виправленого шаблону й наповнює її значеннями."""
    from realtify.paths import PROJECT_ROOT
    from realtify.report_template_import import build_document_from_template

    values = build_report_values(
        intake=intake, task=task, candidates=candidates,
        excel_path=excel_path, excel_values=read_excel_report_values(excel_path),
    )
    template = PROJECT_ROOT / "config" / "report_templates" / "valuation_report_real_template.docx"
    return build_document_from_template(template_path=template, values=values, excel_path=excel_path)


def build_report_document_from_dir(obj_dir: Path) -> dict:
    """Завантажує intake.json / task.generated.yaml / candidates.json / Excel із
    папки об'єкта (web_runs/.../<obj>/) і будує документ."""
    obj_dir = Path(obj_dir)
    intake = _load_intake(_resolve_optional(obj_dir / "intake.json"))
    task = _load_yaml(_resolve_optional(obj_dir / "task.generated.yaml"))
    candidates = _load_candidates(_resolve_optional(obj_dir / "candidates.json"))
    excels = sorted(obj_dir.glob("*_filled.xls")) or sorted(obj_dir.glob("*.xls"))
    excel_path = excels[0] if excels else None
    return build_report_document(intake=intake, task=task, candidates=candidates, excel_path=excel_path)
