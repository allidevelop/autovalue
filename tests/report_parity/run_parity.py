"""Гейт паритету двох рендерерів: один Document JSON → PDF-із-HTML (Playwright) і
PDF-із-docx (mapper → LibreOffice). Растеризуємо обидва (poppler) і рахуємо
попіксельний diff. Запуск:  python tests/report_parity/run_parity.py [obj_dir]
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path

from realtify import report_docx_mapper, report_pdf, report_schema
from realtify.report_document import build_report_document_from_dir


def libreoffice_to_pdf(docx_path: Path, outdir: Path) -> Path:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    subprocess.run(
        [soffice, "--headless", "-env:UserInstallation=file:///tmp/lo_parity",
         "--convert-to", "pdf", "--outdir", str(outdir), str(docx_path)],
        capture_output=True, timeout=180,
    )
    return outdir / (Path(docx_path).stem + ".pdf")


def rasterize(pdf_path: Path, dpi: int = 100):
    from pdf2image import convert_from_path

    from realtify.paths import find_poppler_bin
    poppler = find_poppler_bin()
    kw: dict = {"dpi": dpi}
    if poppler:
        kw["poppler_path"] = str(poppler)
    return convert_from_path(str(pdf_path), **kw)


def diff_ratio(a, b, threshold: int = 24) -> float:
    from PIL import ImageChops
    if a.size != b.size:
        b = b.resize(a.size)
    d = ImageChops.difference(a.convert("L"), b.convert("L"))
    hist = d.histogram()
    total = a.size[0] * a.size[1]
    diff = sum(hist[threshold:])
    return diff / total if total else 1.0


def main() -> int:
    obj_dir = Path(sys.argv[1] if len(sys.argv) > 1 else "/tmp/rost_final5/01_apt_147")
    out = Path("/tmp/parity")
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)

    doc = build_report_document_from_dir(obj_dir)
    errs = report_schema.validate_document(doc)
    print("schema errors:", errs or "none")
    (out / "doc.json").write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8")

    pdf_html = report_pdf.render_document_pdf(doc, out / "from_html.pdf")
    docx = report_docx_mapper.render_document_docx(doc, out / "from_mapper.docx")
    pdf_docx = libreoffice_to_pdf(docx, out)

    imgs_html = rasterize(pdf_html)
    imgs_docx = rasterize(pdf_docx)
    print(f"pages: html={len(imgs_html)} docx={len(imgs_docx)}")
    n = min(len(imgs_html), len(imgs_docx))
    worst = 0.0
    for i in range(n):
        r = diff_ratio(imgs_html[i], imgs_docx[i])
        worst = max(worst, r)
        imgs_html[i].save(out / f"html_{i + 1}.png")
        imgs_docx[i].save(out / f"docx_{i + 1}.png")
        print(f"page {i + 1}: diff={r:.4f}")
    ok = len(imgs_html) == len(imgs_docx) and worst < 0.06
    print(f"WORST diff={worst:.4f}  ->  {'PASS' if ok else 'CHECK'}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
