"""Рендер структурованого документа звіту → HTML (для друку через Playwright page.pdf
та для прев'ю в редакторі). Стилі — зі style-spec; геометрія в мм/pt. Працює тільки
зі словником нод-пересічення CSS∩OOXML, тож docx-рендерер дає той самий вигляд.
"""
from __future__ import annotations

import html
from typing import Any

from realtify import report_styles as styles

_MARK_TAGS = {"bold": ("<strong>", "</strong>"), "italic": ("<em>", "</em>"), "underline": ("<u>", "</u>")}


def _esc(s: Any) -> str:
    return html.escape("" if s is None else str(s))


def css_for_spec(spec: dict[str, Any], mode: str) -> str:
    page = spec["page"]
    m = page["marginMm"]
    body = spec["fonts"]["body"]
    fam = f'"{body["family"]}","{body["fallback"]}",serif'
    p = styles.block_style(spec, "paragraph")
    cap = styles.block_style(spec, "caption")
    show_prov = styles.export_mode(spec, mode).get("showProvenance", False)

    def h(name: str) -> str:
        b = styles.block_style(spec, name)
        return (f".{name}{{font-size:{b['sizePt']}pt;font-weight:{'bold' if b.get('bold') else 'normal'};"
                f"text-align:{b.get('align', 'left')};margin:{b.get('spaceBeforePt', 0)}pt 0 {b.get('spaceAfterPt', 0)}pt 0;}}")

    prov_css = ""
    if show_prov:
        for src in ("auto", "placeholder", "manual"):
            bg = styles.provenance_color(spec, src).get("bg", "transparent")
            prov_css += f".vf-{src}{{background:{bg};}}"

    return f"""
@page {{ size: {page['widthMm']}mm {page['heightMm']}mm; margin: {m['top']}mm {m['right']}mm {m['bottom']}mm {m['left']}mm; }}
* {{ box-sizing: border-box; }}
html,body {{ margin:0; padding:0; }}
body {{ font-family:{fam}; font-size:{body['sizePt']}pt; color:#000; }}
.report {{ width:{page['textWidthMm']}mm; }}
.heading1 {{}} {h('heading1')}
{h('heading2')}
{h('heading3')}
{h('heading4')}
.para {{ font-size:{p['sizePt']}pt; text-align:{p.get('align', 'justify')}; line-height:{p.get('lineHeight', 1.4)};
         text-indent:{p.get('firstLineIndentMm', 0)}mm; margin:0 0 {p.get('spaceAfterPt', 0)}pt 0; }}
.caption {{ font-size:{cap['sizePt']}pt; text-align:{cap.get('align', 'center')};
            margin:{cap.get('spaceBeforePt', 0)}pt 0 {cap.get('spaceAfterPt', 0)}pt 0; }}
.vf {{ }}
{prov_css}
figure {{ margin:0; }}
figure img {{ display:block; }}
.page-break {{ break-after: page; }}
table.locked {{ border-collapse:collapse; table-layout:fixed; margin:6pt 0; }}
table.locked td, table.locked th {{ border:0.5pt solid #000; vertical-align:middle; word-wrap:break-word; }}
""".strip()


def _table_css(spec: dict[str, Any], kind: str) -> str:
    t = styles.table_style(spec, kind)
    return (f"font-size:{t.get('fontSizePt', 8)}pt;text-align:{t.get('align', 'center')};"
            f"padding:{t.get('cellPaddingMm', 1.0)}mm;")


def _render_text(node: dict) -> str:
    out = _esc(node.get("text", ""))
    for mark in node.get("marks", []) or []:
        mt = mark.get("type")
        if mt in _MARK_TAGS:
            o, c = _MARK_TAGS[mt]
            out = f"{o}{out}{c}"
        elif mt == "link":
            href = _esc((mark.get("attrs") or {}).get("href", ""))
            out = f'<a href="{href}">{out}</a>'
    return out


def _render_variable_field(node: dict) -> str:
    a = node.get("attrs") or {}
    src = a.get("source", "auto")
    return (f'<span class="vf vf-{_esc(src)}" data-field="{_esc(a.get("field", ""))}" '
            f'data-source="{_esc(src)}">{_esc(a.get("value", ""))}</span>')


def _render_inline(nodes: list[dict]) -> str:
    parts: list[str] = []
    for n in nodes or []:
        t = n.get("type")
        if t == "text":
            parts.append(_render_text(n))
        elif t == "variableField":
            parts.append(_render_variable_field(n))
        elif t == "image":
            parts.append(_render_image(n))
        else:
            parts.append(_render_inline(n.get("content", [])))
    return "".join(parts)


def _render_image(node: dict) -> str:
    a = node.get("attrs") or {}
    src = _esc(a.get("srcRef", ""))
    width_mm = styles.units.emu_to_mm(a.get("widthEmu", 0)) if a.get("widthEmu") else None
    style = f'width:{width_mm:.2f}mm;' if width_mm else "max-width:100%;"
    cap = a.get("caption")
    href = a.get("href")
    img = f'<img src="{src}" style="{style}">'
    caphtml = f'<figcaption class="caption">{_esc(cap)}</figcaption>' if cap else ""
    if href:
        caphtml = f'<figcaption class="caption"><a href="{_esc(href)}">{_esc(cap or href)}</a></figcaption>'
    return f'<figure>{img}{caphtml}</figure>'


def _render_table(node: dict, spec: dict[str, Any]) -> str:
    a = node.get("attrs") or {}
    kind = a.get("kind", "")
    cols = a.get("columnsMm") or styles.table_style(spec, kind).get("columnsMm", [])
    total = sum(cols) if cols else 0
    cell_css = _table_css(spec, kind)
    bold = styles.table_style(spec, kind).get("headerBold", True)
    colgroup = "".join(f'<col style="width:{w}mm">' for w in cols)
    thead = ""
    if a.get("header"):
        cells = "".join(f'<th style="{cell_css}{"font-weight:bold;" if bold else ""}">{_esc(c)}</th>' for c in a["header"])
        thead = f"<thead><tr>{cells}</tr></thead>"
    body_rows = []
    for row in a.get("rows", []):
        cells = "".join(f'<td style="{cell_css}">{_esc(c)}</td>' for c in row)
        body_rows.append(f"<tr>{cells}</tr>")
    return (f'<table class="locked tbl-{_esc(kind)}" style="width:{total}mm">'
            f'<colgroup>{colgroup}</colgroup>{thead}<tbody>{"".join(body_rows)}</tbody></table>')


def _render_block(node: dict, spec: dict[str, Any]) -> str:
    t = node.get("type")
    if t == "heading":
        lvl = (node.get("attrs") or {}).get("level", 2)
        num = (node.get("attrs") or {}).get("numbering")
        prefix = f"{_esc(num)} " if num else ""
        return f'<h{lvl} class="heading{lvl}">{prefix}{_render_inline(node.get("content", []))}</h{lvl}>'
    if t == "paragraph":
        align = (node.get("attrs") or {}).get("align")
        style = f' style="text-align:{_esc(align)}"' if align else ""
        return f'<p class="para"{style}>{_render_inline(node.get("content", []))}</p>'
    if t == "definitionItem":
        term = _esc((node.get("attrs") or {}).get("term", ""))
        return f'<p class="para"><strong>{term}</strong> {_render_inline(node.get("content", []))}</p>'
    if t in ("bulletList", "orderedList"):
        tag = "ul" if t == "bulletList" else "ol"
        items = "".join(f"<li>{_render_inline(li.get('content', []))}</li>" for li in node.get("content", []))
        return f"<{tag}>{items}</{tag}>"
    if t == "table":
        return _render_table(node, spec)
    if t in ("image", "documentScan"):
        return _render_image(node)
    if t == "pageBreak":
        return '<div class="page-break"></div>'
    if t == "horizontalRule":
        return "<hr>"
    return f'<div>{_render_inline(node.get("content", []))}</div>'


def render_document_html(document: dict, spec: dict[str, Any] | None = None, mode: str = "clean") -> str:
    spec = spec or styles.load_style_spec()
    body = "".join(_render_block(n, spec) for n in document.get("content", []))
    return (f'<!DOCTYPE html><html lang="uk"><head><meta charset="utf-8">'
            f'<style>{css_for_spec(spec, mode)}</style></head>'
            f'<body><article class="report">{body}</article></body></html>')
