from __future__ import annotations

import os
from pathlib import Path

import pytesseract
from PIL import Image

from realtify.paths import TESSDATA_DIR, find_tesseract


def configure_tesseract() -> None:
    tesseract = find_tesseract()
    if tesseract:
        pytesseract.pytesseract.tesseract_cmd = str(tesseract)
    os.environ["TESSDATA_PREFIX"] = str(TESSDATA_DIR)


def ocr_image(image_path: Path, *, lang: str = "ukr+rus+eng") -> str:
    configure_tesseract()
    config = f"--tessdata-dir {TESSDATA_DIR} --psm 6"
    with Image.open(image_path) as image:
        return pytesseract.image_to_string(image, lang=lang, config=config).strip()
