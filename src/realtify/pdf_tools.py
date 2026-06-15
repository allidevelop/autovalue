from __future__ import annotations

from pathlib import Path

from pdf2image import convert_from_path
from pypdf import PdfReader

from realtify.paths import find_poppler_bin


def extract_text_layer(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))
    parts: list[str] = []
    for page in reader.pages:
        parts.append(page.extract_text() or "")
    return "\n".join(parts).strip()


def render_pdf_pages(
    pdf_path: Path,
    output_dir: Path,
    *,
    dpi: int = 220,
    first_page: int | None = None,
    last_page: int | None = None,
) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    poppler_bin = find_poppler_bin()
    pages = convert_from_path(
        str(pdf_path),
        dpi=dpi,
        first_page=first_page,
        last_page=last_page,
        poppler_path=str(poppler_bin) if poppler_bin else None,
    )
    paths: list[Path] = []
    start = first_page or 1
    for offset, image in enumerate(pages):
        path = output_dir / f"page_{start + offset:03d}.png"
        image.save(path)
        paths.append(path)
    return paths

