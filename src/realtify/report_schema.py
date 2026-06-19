"""Схема структурованого документа звіту (ProseMirror/TipTap-сумісний JSON).

Документ = дерево типізованих нод. Кожне динамічне значення — інлайн-нода
`variableField` з {field, source, value}. Розрахункові таблиці/сканы — locked-atom
ноди (`table`/`documentScan`), що НЕ редагуються вручну й пересобираются з даних.

Цей модуль дає (1) білдери нод, (2) валідатор вхідного (відредагованого) JSON,
(3) утиліти обходу. Один словник нод — спільний контракт для HTML- та docx-рендерерів.
"""
from __future__ import annotations

from typing import Any, Iterator

# ── типи нод/marks ───────────────────────────────────────────────────────────
BLOCK_TYPES = {
    "heading", "paragraph", "bulletList", "orderedList", "listItem",
    "definitionItem", "table", "documentScan", "pageBreak", "signatureBlock",
    "horizontalRule",
}
INLINE_TYPES = {"text", "variableField", "image"}
MARK_TYPES = {"bold", "italic", "underline", "link"}
LOCKED_TYPES = {"table", "documentScan"}  # не редагуються вручну
SOURCES = {"auto", "placeholder", "manual"}

# Маркери ручного заповнення (узгоджено з report_generator._FILL_BLANK).
_FILL_MARKERS = ("________", "заповнити")


def classify_source(value: Any) -> str:
    """auto vs placeholder за вмістом значення (manual виставляє лише UI при правці)."""
    s = "" if value is None else str(value)
    low = s.lower()
    if any(m in low for m in _FILL_MARKERS) or not s.strip():
        return "placeholder"
    return "auto"


# ── білдери ──────────────────────────────────────────────────────────────────
def doc(content: list[dict]) -> dict:
    return {"type": "doc", "version": 1, "content": content}


def heading(level: int, content: list[dict], numbering: str | None = None) -> dict:
    attrs: dict[str, Any] = {"level": int(level)}
    if numbering:
        attrs["numbering"] = numbering
    return {"type": "heading", "attrs": attrs, "content": content}


def paragraph(content: list[dict], align: str | None = None) -> dict:
    node: dict[str, Any] = {"type": "paragraph", "content": content}
    if align:
        node["attrs"] = {"align": align}
    return node


def text(s: str, marks: list[dict] | None = None) -> dict:
    node: dict[str, Any] = {"type": "text", "text": s}
    if marks:
        node["marks"] = marks
    return node


def variable_field(field: str, value: Any, source: str | None = None, fmt: str | None = None) -> dict:
    val = "" if value is None else str(value)
    attrs: dict[str, Any] = {"field": field, "value": val, "source": source or classify_source(val)}
    if fmt:
        attrs["format"] = fmt
    return {"type": "variableField", "attrs": attrs}


def table(kind: str, header: list[str], rows: list[list[str]], columns_mm: list[float]) -> dict:
    return {
        "type": "table",
        "attrs": {
            "kind": kind, "header": list(header), "rows": [list(r) for r in rows],
            "columnsMm": list(columns_mm), "source": "excel_sidecar", "locked": True,
        },
    }


def page_break() -> dict:
    return {"type": "pageBreak"}


# marks
def bold() -> dict:
    return {"type": "bold"}


def italic() -> dict:
    return {"type": "italic"}


def underline() -> dict:
    return {"type": "underline"}


def link(href: str) -> dict:
    return {"type": "link", "attrs": {"href": href}}


# ── обхід ─────────────────────────────────────────────────────────────────────
def iter_nodes(node: dict) -> Iterator[dict]:
    yield node
    for child in node.get("content", []) or []:
        yield from iter_nodes(child)


def iter_variable_fields(document: dict) -> Iterator[dict]:
    for n in iter_nodes(document):
        if n.get("type") == "variableField":
            yield n


def unfilled_placeholders(document: dict) -> list[str]:
    """Поля, що лишились placeholder (для блокування clean-експорту)."""
    return [
        n["attrs"].get("field", "?")
        for n in iter_variable_fields(document)
        if n.get("attrs", {}).get("source") == "placeholder"
    ]


# ── валідація ─────────────────────────────────────────────────────────────────
def validate_document(document: Any) -> list[str]:
    """Повертає список помилок (порожній = ок). Перевіряє словник типів і атрибути."""
    errors: list[str] = []
    if not isinstance(document, dict) or document.get("type") != "doc":
        return ["root: expected node of type 'doc'"]
    content = document.get("content")
    if not isinstance(content, list):
        return ["root: 'content' must be a list"]

    def check(node: Any, path: str) -> None:
        if not isinstance(node, dict):
            errors.append(f"{path}: node must be an object")
            return
        t = node.get("type")
        if t == "text":
            if not isinstance(node.get("text"), str):
                errors.append(f"{path}: text node needs string 'text'")
        elif t == "variableField":
            a = node.get("attrs") or {}
            if not a.get("field"):
                errors.append(f"{path}: variableField needs 'field'")
            if a.get("source") not in SOURCES:
                errors.append(f"{path}: variableField source invalid")
        elif t == "table":
            a = node.get("attrs") or {}
            if a.get("kind") not in {"adjustment_73", "comparables_71"}:
                errors.append(f"{path}: table kind invalid")
            if not isinstance(a.get("rows"), list) or not isinstance(a.get("columnsMm"), list):
                errors.append(f"{path}: table needs rows[] and columnsMm[]")
        elif t in BLOCK_TYPES or t in INLINE_TYPES:
            pass
        else:
            errors.append(f"{path}: unknown node type {t!r}")
        for i, child in enumerate(node.get("content", []) or []):
            check(child, f"{path}.{t}[{i}]")
        for m in node.get("marks", []) or []:
            if not isinstance(m, dict) or m.get("type") not in MARK_TYPES:
                errors.append(f"{path}: invalid mark {m!r}")

    for i, child in enumerate(content):
        check(child, f"doc[{i}]")
    return errors
