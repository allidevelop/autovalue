"""Єдина style-spec звіту — один контракт стилів для обох рендерерів
(HTML/PDF та .docx). Геометрія задана в мм/pt; рендерери конвертують однаково,
тож вихідні документи не розходяться.
"""
from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path
from typing import Any

from realtify.report_styles import units  # noqa: F401 (re-export)

_SPEC_PATH = Path(__file__).with_name("style_spec.json")


@lru_cache(maxsize=1)
def load_style_spec() -> dict[str, Any]:
    """Парсить style_spec.json (кешовано)."""
    return json.loads(_SPEC_PATH.read_text(encoding="utf-8"))


def block_style(spec: dict[str, Any], name: str) -> dict[str, Any]:
    return dict(spec.get("blocks", {}).get(name, {}))


def table_style(spec: dict[str, Any], kind: str) -> dict[str, Any]:
    return dict(spec.get("tables", {}).get(kind, {}))


def provenance_color(spec: dict[str, Any], source: str) -> dict[str, Any]:
    return dict(spec.get("provenance", {}).get(source, {}))


def export_mode(spec: dict[str, Any], mode: str) -> dict[str, Any]:
    modes = spec.get("exportModes", {})
    return dict(modes.get(mode, modes.get("clean", {"showProvenance": False})))
