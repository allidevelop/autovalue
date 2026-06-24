"""E2E-самотест коза-клон генерації: будує звіт КЛОНУВАННЯМ кози з реального витяга,
рендерить сторінки й формує HTML-сторінку «ВХІД (з документів) vs ВИХІД (у звіті)».
Призначення — чесна наскрізна перевірка: чи дані з документів реально потрапили у звіт.
"""
from __future__ import annotations

import base64
import glob
import html
import re
import shutil
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path

from realtify.paths import PROJECT_ROOT

E2E_DIR = PROJECT_ROOT / "web_runs" / "_e2e"
E2E_HTML = E2E_DIR / "e2e_report.html"
E2E_STATUS = E2E_DIR / "status.txt"
# Тестовий об'єкт за замовчуванням — найсвіжіший batch-job із вхідним PDF.
_DEFAULT_JOBS_ROOT = PROJECT_ROOT / "web_runs"


def _latest_input_pdf() -> Path | None:
    cands: list[tuple[float, Path]] = []
    for inp in _DEFAULT_JOBS_ROOT.glob("*/input"):
        for pdf in inp.glob("*.pdf"):
            if "merged" not in pdf.name.lower() or len(list(inp.glob("*.pdf"))) == 1:
                cands.append((pdf.stat().st_mtime, pdf))
    return max(cands)[1] if cands else None


def _num(s) -> float | None:
    if not s:
        return None
    try:
        return float(str(s).replace(" ", "").replace(" ", "").replace(",", "."))
    except (TypeError, ValueError):
        return None


def _set_status(text: str) -> None:
    E2E_DIR.mkdir(parents=True, exist_ok=True)
    E2E_STATUS.write_text(text, encoding="utf-8")


def build_e2e_report(pdf_path: Path | None = None, *, dpi: int = 300, stamp: str = "") -> Path:
    """Генерує E2E-звіт (HTML) для одного об'єкта. stamp — мітка часу ззовні."""
    from realtify import koza_engine as ke
    from realtify.intake import extract_intake_from_pdf
    from realtify.report_generator import build_report_values
    from realtify.valuation_date import resolve_valuation_context
    import yaml

    E2E_DIR.mkdir(parents=True, exist_ok=True)
    pdf = pdf_path or _latest_input_pdf()
    if not pdf or not pdf.exists():
        raise RuntimeError("Не знайдено вхідний PDF (витяг) для тесту.")

    _set_status("running")
    work = E2E_DIR / "work"
    shutil.rmtree(work, ignore_errors=True)
    objdir = work / "01_apt"
    objdir.mkdir(parents=True)

    # 1) свіжий intake (як реальний пайплайн)
    intake = extract_intake_from_pdf(pdf_path=pdf, output_dir=objdir, dpi=dpi, profile="apartment")
    task = yaml.safe_load((objdir / "task.generated.yaml").read_text(encoding="utf-8"))
    values = build_report_values(intake=intake.result, task=task, candidates=[],
                                 excel_path=None, excel_values={})
    ctx = resolve_valuation_context(task, excel_path=None)

    # 2) коза-клон
    out = ke.build_report_via_koza(objdir, work)
    if not out:
        _set_status("no-koza")
        raise RuntimeError("Для дому об'єкта немає кози у базі шаблонів — клон не побудовано.")
    koza_name = out.name
    clone_txt = ke._doc_to_text(out)

    se = intake.result.selected_extract

    def out_find(pat, default="(немає)"):
        m = re.search(pat, clone_txt)
        return m.group(1).strip() if m else default

    out_apt = out_find(r"квартир[аи]\s*№\s*(\d+)")
    out_area = out_find(r"загальною площею\s*([\d.,]+)\s*кв")
    out_rooms = out_find(r"(одно|дво|три|чотири|п['’]?яти|шести)кімнатн")
    out_dov = out_find(r"Дата оцінки:\s*([^\n]+)")
    out_dsk = out_find(r"Дата складання звіту:\s*([^\n]+)")
    out_val = out_find(r"Ринкова вартість[^\d(]*([\d  ]+[,.]\d{2})")
    out_words = out_find(r"Ринкова вартість[^()]*\(([^)]+)\)")
    out_ref = out_find(r"реєстрацію прав власності\s*(№[^\n]{0,40})", "(прибрано — на скані)")

    # порівняння
    rooms_map = {1: "одно", 2: "дво", 3: "три", 4: "чотири", 5: "п'яти", 6: "шести"}
    exp_rooms = rooms_map.get(values.get("rooms_count")) if values.get("rooms_count") else None
    rows = [
        ("Адреса", se.address_full if se else "", out_find(r"за адресою:\s*([^\n]+)"), None),
        ("№ квартири", str(se.apartment_number if se else ""), out_apt,
         str(se.apartment_number if se else "") == out_apt),
        ("Площа загальна, кв.м", _fmt(se.total_area_m2 if se else None),
         out_area, _num(se.total_area_m2 if se else None) == _num(out_area)),
        ("Площа житлова, кв.м", _fmt(se.living_area_m2 if se else None), "—", None),
        ("Кімнатність", str(values.get("rooms_count") or "(не розпізнано)"),
         out_rooms + "кімнатна" if out_rooms != "(немає)" else out_rooms,
         (exp_rooms == out_rooms) if exp_rooms else None),
        ("Дата оцінки (з Excel-реєстру)", ctx.valuation_date.strftime("%d.%m.%Y"), out_dov,
         _date_ok(ctx.valuation_date, out_dov)),
        ("Дата складання звіту", ctx.valuation_date.strftime("%d.%m.%Y"), out_dsk,
         _date_ok(ctx.valuation_date, out_dsk)),
        ("Вартість (число)", "≈ ціна/м² кози × площа", out_val, None),
        ("Вартість (прописом)", "узгоджено з числом", out_words, None),
        ("№ витяга", values.get("extract_index_number") or "(OCR не розпізнав)", out_ref, None),
    ]

    # 3) рендер сторінок
    pdf_out = ke._soffice_convert(out, work, "pdf", ".pdf")
    imgs64: list[str] = []
    if pdf_out and pdf_out.exists():
        subprocess.run(["pdftoppm", "-png", "-r", "110", "-f", "1", "-l", "22",
                        str(pdf_out), str(work / "page")], capture_output=True, timeout=240)
        for p in sorted(glob.glob(str(work / "page-*.png"))):
            imgs64.append("data:image/png;base64," + base64.b64encode(Path(p).read_bytes()).decode())

    _write_html(rows, imgs64, koza_name, pdf.name, stamp or _now())
    _set_status("ready")
    return E2E_HTML


def _fmt(v):
    if v in (None, ""):
        return ""
    s = f"{float(v):.2f}".rstrip("0").rstrip(".")
    return s.replace(".", ",")


def _date_ok(d, out_text):
    months = ["січня", "лютого", "березня", "квітня", "травня", "червня",
              "липня", "серпня", "вересня", "жовтня", "листопада", "грудня"]
    want = f"{d.day:02d} {months[d.month - 1]} {d.year}"
    want2 = f"{d.day} {months[d.month - 1]} {d.year}"
    return (want in out_text) or (want2 in out_text)


def _now():
    return ""


def _write_html(rows, imgs64, koza_name, pdf_name, stamp):
    def cell(v):
        return html.escape(str(v)) if v not in (None, "") else "—"

    trs = ""
    for name, iv, ov, ok in rows:
        badge = "" if ok is None else ('<span class=ok>✓</span>' if ok else '<span class=bad>✗</span>')
        trs += (f"<tr><td>{html.escape(name)}</td><td class=in>{cell(iv)}</td>"
                f"<td class=out>{cell(ov)} {badge}</td></tr>")
    imgs = "".join(f'<div class=pg><img src="{u}"></div>' for u in imgs64)
    page = f"""<!doctype html><html lang=uk><head><meta charset=utf-8>
<meta name=viewport content="width=device-width, initial-scale=1">
<title>E2E коза-клон: вхід → вихід</title><style>
body{{font-family:system-ui,Segoe UI,Roboto,sans-serif;margin:24px;color:#111;background:#fafafa}}
h1{{font-size:20px;margin:0 0 4px}} .meta{{color:#666;font-size:13px;margin-bottom:16px}}
table{{border-collapse:collapse;width:100%;max-width:1100px;background:#fff}}
td,th{{border:1px solid #d7d7d7;padding:9px 11px;vertical-align:top;font-size:14px}}
th{{background:#f1f3f5;text-align:left}} td:first-child{{font-weight:600;white-space:nowrap}}
.in{{background:#eef5ff}} .out{{background:#eefbf1}}
.ok{{color:#16794c;font-weight:700}} .bad{{color:#c0341d;font-weight:700}}
h2{{font-size:16px;margin:26px 0 10px}}
.pgs{{display:flex;flex-wrap:wrap;gap:10px}} .pg{{border:1px solid #ddd;background:#fff}} .pg img{{width:300px;display:block}}
</style></head><body>
<h1>E2E-самотест: коза-клон генерації звіту</h1>
<div class=meta>Вхід — дані, розпізнані з документів (витяг/техпаспорт, 300&nbsp;dpi) та з Excel-реєстру дат.
Вихід — що реально потрапило у згенерований звіт (клон шаблону дому).<br>
Коза: <b>{html.escape(koza_name)}</b> · вхідний PDF: {html.escape(pdf_name)} · згенеровано: {html.escape(stamp)}</div>
<table><tr><th>Поле</th><th>ВХІД (з документів)</th><th>ВИХІД (у звіті)</th></tr>{trs}</table>
<h2>Сторінки згенерованого звіту</h2><div class=pgs>{imgs}</div>
</body></html>"""
    E2E_DIR.mkdir(parents=True, exist_ok=True)
    E2E_HTML.write_text(page, encoding="utf-8")
