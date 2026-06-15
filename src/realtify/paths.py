from __future__ import annotations

import sys
import os
import shutil
from pathlib import Path


def _project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def _resource_root() -> Path:
    bundle_root = getattr(sys, "_MEIPASS", None)
    if bundle_root:
        return Path(bundle_root).resolve()
    return Path(__file__).resolve().parents[2]


PROJECT_ROOT = _project_root()
RESOURCE_ROOT = _resource_root()
TOOLS_DIR = RESOURCE_ROOT / "tools"
TESSDATA_DIR = TOOLS_DIR / "tessdata"
POPPLER_ROOT = TOOLS_DIR / "poppler"
PLAYWRIGHT_BROWSERS_ROOT = TOOLS_DIR / "ms-playwright"

if PLAYWRIGHT_BROWSERS_ROOT.exists():
    os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(PLAYWRIGHT_BROWSERS_ROOT))


def find_tesseract() -> Path | None:
    candidates = [
        Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
        Path.home() / r"AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    found = shutil.which("tesseract")
    if found:
        return Path(found)
    return None


def find_poppler_bin() -> Path | None:
    for candidate in POPPLER_ROOT.rglob("pdfinfo.exe"):
        return candidate.parent
    found = shutil.which("pdfinfo")
    if found:
        return Path(found).parent
    return None


def find_tessdata_dir() -> Path | None:
    env_value = os.environ.get("TESSDATA_PREFIX")
    candidates: list[Path] = []
    if env_value:
        env_path = Path(env_value)
        candidates.extend([env_path, env_path / "tessdata"])
    candidates.extend(
        [
            TESSDATA_DIR,
            Path(r"C:\Program Files\Tesseract-OCR\tessdata"),
            Path("/usr/share/tesseract-ocr/5/tessdata"),
            Path("/usr/share/tesseract-ocr/4.00/tessdata"),
            Path("/usr/share/tessdata"),
            Path("/usr/local/share/tessdata"),
        ]
    )
    for candidate in candidates:
        if _valid_tessdata_dir(candidate):
            return candidate
    return None


def _valid_tessdata_dir(path: Path) -> bool:
    return path.exists() and any((path / f"{lang}.traineddata").exists() for lang in ("ukr", "rus", "eng"))


def ensure_output_dir(task_id: str) -> Path:
    out = PROJECT_ROOT / "outputs" / task_id
    out.mkdir(parents=True, exist_ok=True)
    return out
