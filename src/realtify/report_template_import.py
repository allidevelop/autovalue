"""Імпорт структури вже-виправленого .docx-шаблону у schema-ноди (Фаза 2).

Проходить тіло шаблону по порядку (абзаци + таблиці), переиспользует весь
бойлерплейт і {{плейсхолдери}}:
- абзац → heading/paragraph; рядки → text+marks; {{key}} → variableField(values[key]);
- таблиця поправок (27×8) → locked-нода із сайдкара;
- таблиця порівняння (11×7, у шаблоні — зразкові дані) → locked-нода з даних аналогів;
- інші таблиці (статика без витоків) → locked-снапшот із підстановкою {{}};
- картинки (16) — Фаза 2b (поки пропускаємо).
"""
from __future__ import annotations

import io
import re
import tempfile
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from realtify import report_schema as S
from realtify import report_styles as styles
from realtify.report_document import _ADJ_HEADER, _ADJ_ROW_ORDER  # reuse
from realtify.report_scans import candidate_image_bytes, render_object_scans, template_media_bytes, to_data_uri
from realtify.excel_sidecar import sidecar_adjustment_rows

# Слоти аналогів у шаблоні (порядок документа = аналог 1→5).
_ANALOG_SLOTS = ["image14.png", "image13.png", "image11.png", "image12.png", "image9.png"]

_PH = re.compile(r"\{\{(\w+)\}\}")
_TWIPS_PER_MM = 56.6929

# Рядки таблиці порівняння (label, comparable_N_<key>, object_value_key|None).
_CMP_ROWS = [
    ("Адреса об'єкта порівняння", "address", "address_short"),
    ("Площа, кв.м", "area_m2", "total_area_m2"),
    ("Поверховість", "floor_or_level", "floor_or_level"),
    ("Ціна пропозиції, $ США", "price_usd", None),
    ("Ціна 1 м², $ США", "price_usd_m2", None),
    ("Місце розташування", "location_quality", None),
    ("Клас житлового комплексу", "building_class", None),
    ("Оздоблення", "condition", None),
    ("Термін здачі", "delivery_date", None),
    ("Джерело інформації", "source_url", None),
]


def _iter_body(doc):
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            yield Paragraph(child, doc)
        elif tag == "tbl":
            yield Table(child, doc)


def _is_heading(p: Paragraph) -> bool:
    t = p.text.strip()
    runs = [r for r in p.runs if r.text.strip()]
    all_bold = bool(runs) and all(r.bold for r in runs)
    return bool(t and len(t) < 90 and all_bold and (re.match(r"^\d+[.\s]", t) or t == t.upper()))


def _heading_level(p: Paragraph) -> int:
    t = p.text.strip()
    if re.match(r"^\d+\.\d", t):
        return 3
    if re.match(r"^\d+\.", t):
        return 2
    return 1


def _inline(p: Paragraph, values: dict[str, Any]) -> list[dict]:
    out: list[dict] = []
    for run in p.runs:
        marks = []
        if run.bold:
            marks.append(S.bold())
        if run.italic:
            marks.append(S.italic())
        if run.underline:
            marks.append(S.underline())
        txt = run.text
        last = 0
        for m in _PH.finditer(txt):
            if m.start() > last:
                out.append(S.text(txt[last:m.start()], marks or None))
            key = m.group(1)
            out.append(S.variable_field(key, values.get(key, "")))
            last = m.end()
        if last < len(txt):
            out.append(S.text(txt[last:], marks or None))
    return out


def _subst(text: str, values: dict[str, Any]) -> str:
    return _PH.sub(lambda m: str(values.get(m.group(1), "")), text)


def _col_widths_mm(table: Table, text_width_mm: float) -> list[float]:
    grid = table._tbl.tblGrid
    widths = []
    for c in grid.findall(qn("w:gridCol")):
        w = c.get(qn("w:w"))
        widths.append(float(w) / _TWIPS_PER_MM if w else 0.0)
    n = len(table.columns)
    total = sum(widths)
    if total <= 0 or len(widths) != n:
        return [text_width_mm / n] * n
    return [w / total * text_width_mm for w in widths]


def _adjustment_node(excel_path: Path | None, spec: dict[str, Any]) -> dict:
    rows_map = sidecar_adjustment_rows(excel_path) if excel_path else {}
    rows = []
    for idx in _ADJ_ROW_ORDER:
        r = rows_map.get(idx) or rows_map.get(str(idx))
        if r:
            rows.append([str(c) for c in r[:8]])
    return S.table("adjustment_73", _ADJ_HEADER, rows, styles.table_style(spec, "adjustment_73")["columnsMm"])


def _comparables_node(values: dict[str, Any], table: Table, spec: dict[str, Any]) -> dict:
    header = ["Показник", "Аналог №1", "Аналог №2", "Аналог №3", "Аналог №4", "Аналог №5", "Об'єкт оцінки"]
    rows = []
    for label, ckey, okey in _CMP_ROWS:
        row = [label]
        for n in range(1, 6):
            row.append(str(values.get(f"comparable_{n}_{ckey}", "")))
        row.append(str(values.get(okey, "")) if okey else ("Х" if ckey == "source_url" else ""))
        rows.append(row)
    cols = _col_widths_mm(table, spec["page"]["textWidthMm"])
    return S.table("comparables_71", header, rows, cols)


def _generic_table_node(table: Table, values: dict[str, Any], spec: dict[str, Any]) -> dict:
    rows = [[_subst(c.text.strip(), values) for c in row.cells] for row in table.rows]
    if not any(any(cell for cell in r) for r in rows):
        return {}
    cols = _col_widths_mm(table, spec["page"]["textWidthMm"])
    return S.table("generic", [], rows, cols)


def _table_has_fill_markers(table: Table) -> bool:
    for row in table.rows:
        for c in row.cells:
            t = c.text
            if "заповнити" in t.lower() or "________" in t:
                return True
    return False


def _editable_table_blocks(table: Table, ti: int) -> list[dict]:
    """Таблиця з ручними плейсхолдерами (напр. «Характеристика місця розташування»)
    → редаговані рядки «Мітка: [поле]» (замість read-only снапшота)."""
    rows = list(table.rows)
    out: list[dict] = []
    if rows:
        hdr = rows[0].cells[0].text.strip()
        if hdr:
            out.append(S.heading(4, [S.text(hdr)]))
    for ri in range(1, len(rows)):
        cells = rows[ri].cells
        label = " ".join(cells[0].text.split()).rstrip(":")
        if len(cells) >= 2 and cells[0]._tc is not cells[1]._tc:
            value = cells[1].text.strip()
            out.append(S.paragraph([
                S.text(label + ": ", [S.bold()]),
                S.variable_field(f"char_{ti}_{ri}", value or "________________", label=label),
            ]))
        elif label:
            out.append(S.paragraph([S.text(label)]))
    return out


def _is_comparables(table: Table) -> bool:
    if len(table.columns) != 7 or len(table.rows) < 6:
        return False
    txt = " ".join(c.text for c in table.rows[0].cells).lower()
    col0 = " ".join(table.rows[r].cells[0].text for r in range(len(table.rows))).lower()
    return "об'єкт порівняння" in txt and "адреса об'єкта порівняння" in col0


def _aspect_bytes(b: bytes) -> float:
    try:
        from PIL import Image
        with Image.open(io.BytesIO(b)) as im:
            return im.height / im.width if im.width else 1.414
    except Exception:  # noqa: BLE001
        return 1.414


def _fit_width(width_emu: int, aspect: float, max_h: int) -> int:
    return round(max_h / aspect) if width_emu * aspect > max_h else width_emu


def _drawings_in(el, doc) -> list[tuple[str, int, int]]:
    res = []
    for dr in el.findall(".//" + qn("w:drawing")):
        ext = dr.find(".//" + qn("wp:extent"))
        cx = int(ext.get("cx")) if ext is not None and ext.get("cx") else 0
        cy = int(ext.get("cy")) if ext is not None and ext.get("cy") else 0
        blip = dr.find(".//" + qn("a:blip"))
        rid = blip.get(qn("r:embed")) if blip is not None else None
        name = ""
        if rid and rid in doc.part.rels:
            try:
                name = doc.part.rels[rid].target_ref.split("/")[-1]
            except Exception:  # noqa: BLE001
                name = ""
        if name:
            res.append((name, cx, cy))
    return res


def _build_image_map(intake, task, candidates, tmp_dir, spec) -> dict[str, tuple]:
    """media_name → (type, kind, data_uri, width_emu, aspect, caption, href)."""
    full_w = spec["images"]["fullWidthEmu"]
    max_h = spec["images"]["maxHeightEmu"]
    m: dict[str, tuple] = {}
    scans = render_object_scans(intake, task, tmp_dir)
    for slot, kind in (("image5.png", "vityag"), ("image1.png", "techpass")):
        if kind in scans:
            b, asp = scans[kind]
            m[slot] = ("documentScan", kind, to_data_uri(b), _fit_width(full_w, asp, max_h), asp, None, None)
    for i, slot in enumerate(_ANALOG_SLOTS):
        if i < len(candidates):
            b = candidate_image_bytes(candidates[i])
            if b:
                asp = _aspect_bytes(b)
                cand = candidates[i]
                cap = (getattr(cand, "address", None) or str(getattr(cand, "source_url", "")) or "").strip()
                href = str(getattr(cand, "source_url", "")) or None
                m[slot] = ("image", None, to_data_uri(b), _fit_width(full_w, asp, max_h), asp, cap, href)
    return m


def build_document_from_template(
    *, template_path: Path, values: dict[str, Any], excel_path: Path | None,
    intake=None, task: dict | None = None, candidates: list | None = None,
) -> dict:
    spec = styles.load_style_spec()
    doc = Document(str(template_path))
    from realtify.report_generator import _is_existing_adjustment_table
    candidates = candidates or []

    content: list[dict] = []
    ti = 0
    with tempfile.TemporaryDirectory(prefix="rep_img_") as tmp:
        img_map = _build_image_map(intake, task, candidates, Path(tmp), spec)
        for block in _iter_body(doc):
            if isinstance(block, Paragraph):
                draws = _drawings_in(block._p, doc)
                if draws:
                    for name, cx, cy in draws:
                        entry = img_map.get(name)
                        if entry:
                            typ, kind, uri, w, asp, cap, href = entry
                            if typ == "documentScan":
                                content.append(S.document_scan(kind, uri, w, asp))
                            else:
                                content.append(S.image(uri, w, asp, caption=cap, href=href))
                        else:  # статична картинка шаблону (штамп/підпис/сертифікат/лого)
                            sb = template_media_bytes(template_path, name)
                            if sb and cx:
                                content.append(S.image(to_data_uri(sb), cx, (cy / cx) if cx else 1.0))
                    continue
                inline = _inline(block, values)
                if not any((n.get("text", "").strip() if n["type"] == "text" else True) for n in inline):
                    continue
                if _is_heading(block):
                    content.append(S.heading(_heading_level(block), inline))
                else:
                    content.append(S.paragraph(inline))
            else:  # Table
                ti += 1
                if _is_existing_adjustment_table(block):
                    content.append(_adjustment_node(excel_path, spec))
                elif _is_comparables(block):
                    content.append(_comparables_node(values, block, spec))
                elif _table_has_fill_markers(block):
                    content.extend(_editable_table_blocks(block, ti))
                else:
                    node = _generic_table_node(block, values, spec)
                    if node:
                        content.append(node)
                # Сканы об'єкта (витяг/техпаспорт) у клітинках таблиці → documentScan.
                for name, _cx, _cy in _drawings_in(block._tbl, doc):
                    entry = img_map.get(name)
                    if entry and entry[0] == "documentScan":
                        content.append(S.document_scan(entry[1], entry[2], entry[3], entry[4]))
    return S.doc(content)
