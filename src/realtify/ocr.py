from __future__ import annotations

import os
from pathlib import Path

import pytesseract
from PIL import Image

from realtify.paths import find_tessdata_dir, find_tesseract


DEFAULT_OCR_TIMEOUT_SECONDS = 45
DEFAULT_OCR_PSM = 11


class OcrTimeoutError(RuntimeError):
    pass


def configure_tesseract() -> Path | None:
    tesseract = find_tesseract()
    if tesseract:
        pytesseract.pytesseract.tesseract_cmd = str(tesseract)
    tessdata_dir = find_tessdata_dir()
    if tessdata_dir:
        os.environ["TESSDATA_PREFIX"] = str(tessdata_dir)
    return tessdata_dir


def ocr_image(image_path: Path, *, lang: str = "ukr+rus+eng", timeout_seconds: int | None = None) -> str:
    tessdata_dir = configure_tesseract()
    timeout = timeout_seconds or _ocr_timeout_seconds()
    config_parts = []
    if tessdata_dir:
        config_parts.append(f'--tessdata-dir "{tessdata_dir}"')
    config_parts.extend(["--psm", str(_ocr_psm())])
    with Image.open(image_path) as image:
        try:
            return pytesseract.image_to_string(
                image,
                lang=lang,
                config=" ".join(config_parts),
                timeout=timeout,
            ).strip()
        except RuntimeError as exc:
            if "timeout" in str(exc).casefold():
                raise OcrTimeoutError(f"OCR timeout after {timeout}s for {image_path}") from exc
            raise


def _ocr_timeout_seconds() -> int:
    raw_value = os.getenv("REALTIFY_OCR_TIMEOUT_SECONDS")
    if raw_value:
        try:
            value = int(raw_value)
            if value > 0:
                return value
        except ValueError:
            pass
    return DEFAULT_OCR_TIMEOUT_SECONDS


def _ocr_psm() -> int:
    raw_value = os.getenv("REALTIFY_OCR_PSM")
    if raw_value:
        try:
            value = int(raw_value)
            if value > 0:
                return value
        except ValueError:
            pass
    return DEFAULT_OCR_PSM
