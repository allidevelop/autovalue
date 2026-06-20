"""Зображення для звіту (Фаза 2b): рендер сканів об'єкта (витяг/техпаспорт) із
вихідного PDF та витяг статичних картинок шаблону. Переиспользует хелпери
report_generator/pdf_tools, але повертає БАЙТИ (для вбудовування як data-URI у
schema-документ, що робить JSON самодостатнім).
"""
from __future__ import annotations

import base64
import zipfile
from pathlib import Path
from typing import Any


def to_data_uri(png_bytes: bytes) -> str:
    return "data:image/png;base64," + base64.b64encode(png_bytes).decode("ascii")


def render_object_scans(intake: Any, task: dict | None, tmp_dir: Path) -> dict[str, tuple[bytes, float]]:
    """{'vityag': (png, aspect), 'techpass': (png, aspect)} — сторінки об'єкта з PDF."""
    from realtify.report_generator import _image_aspect, _pick_techpass_pages, _source_pdf, _stack_images_vertical
    from realtify.pdf_tools import render_pdf_pages

    out: dict[str, tuple[bytes, float]] = {}
    if intake is None:
        return out
    src = _source_pdf(intake, task or {})
    if not src or not Path(src).exists():
        return out

    se = getattr(intake, "selected_extract", None)
    st = getattr(intake, "selected_technical_passport", None)

    vp = getattr(se, "page", None) if se else None
    if vp:
        try:
            imgs = render_pdf_pages(src, tmp_dir / "v", first_page=vp, last_page=vp, dpi=160)
            if imgs:
                p = Path(imgs[0])
                out["vityag"] = (p.read_bytes(), _image_aspect(p))
        except Exception:  # noqa: BLE001
            pass

    tp_all = list(getattr(st, "pages", []) or []) if st else []
    tp = _pick_techpass_pages(src, tp_all, ocr_dir=getattr(intake, "pages_text_dir", None)) if tp_all else []
    rendered: list[Path] = []
    for i, page in enumerate(tp):
        try:
            imgs = render_pdf_pages(src, tmp_dir / f"t{i}", first_page=page, last_page=page, dpi=160)
            if imgs:
                rendered.append(Path(imgs[0]))
        except Exception:  # noqa: BLE001
            pass
    if len(rendered) == 1:
        out["techpass"] = (rendered[0].read_bytes(), _image_aspect(rendered[0]))
    elif len(rendered) > 1:
        try:
            stacked, aspect = _stack_images_vertical(rendered, tmp_dir / "tp_stack.png")
            out["techpass"] = (Path(stacked).read_bytes(), aspect)
        except Exception:  # noqa: BLE001
            out["techpass"] = (rendered[0].read_bytes(), _image_aspect(rendered[0]))
    return out


def template_media_bytes(template_path: Path, media_name: str) -> bytes | None:
    """Байти статичної картинки шаблону (штампи/підписи/сертифікати/лого)."""
    try:
        with zipfile.ZipFile(template_path) as z:
            return z.read(f"word/media/{media_name}")
    except (KeyError, zipfile.BadZipFile, FileNotFoundError):
        return None


def candidate_image_bytes(candidate: Any) -> bytes | None:
    img = getattr(candidate, "report_image_path", None) or getattr(candidate, "screenshot_path", None)
    if img and Path(str(img)).exists():
        try:
            return Path(str(img)).read_bytes()
        except Exception:  # noqa: BLE001
            return None
    return None
