from __future__ import annotations

import os
from pathlib import Path

import pytesseract
from PIL import Image

from realtify.paths import find_tessdata_dir, find_tesseract


def configure_tesseract() -> Path | None:
    tesseract = find_tesseract()
    if tesseract:
        pytesseract.pytesseract.tesseract_cmd = str(tesseract)
    tessdata_dir = find_tessdata_dir()
    if tessdata_dir:
        os.environ["TESSDATA_PREFIX"] = str(tessdata_dir)
    return tessdata_dir


def ocr_image(image_path: Path, *, lang: str = "ukr+rus+eng") -> str:
    tessdata_dir = configure_tesseract()
    config_parts = []
    if tessdata_dir:
        config_parts.append(f'--tessdata-dir "{tessdata_dir}"')
    config_parts.extend(["--psm", "6"])
    with Image.open(image_path) as image:
        return pytesseract.image_to_string(image, lang=lang, config=" ".join(config_parts)).strip()
