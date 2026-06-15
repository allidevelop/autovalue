from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path

from rich.console import Console
from rich.table import Table

from realtify.paths import PLAYWRIGHT_BROWSERS_ROOT, PROJECT_ROOT, find_poppler_bin, find_tessdata_dir, find_tesseract


REQUIRED_MODULES = [
    "bs4",
    "lxml",
    "docx",
    "olefile",
    "pdf2image",
    "PIL",
    "playwright",
    "pydantic",
    "pypdf",
    "pytesseract",
    "win32com.client",
    "yaml",
    "requests",
    "rich",
]


def _module_ok(name: str) -> bool:
    return importlib.util.find_spec(name) is not None


def _command_output(args: list[str]) -> str:
    try:
        result = subprocess.run(
            args,
            capture_output=True,
            text=True,
            check=False,
            timeout=15,
        )
    except Exception as exc:
        return f"error: {exc}"
    output = (result.stdout or result.stderr).strip()
    return output.splitlines()[0] if output else f"exit {result.returncode}"


def collect_status() -> dict[str, object]:
    tesseract = find_tesseract()
    poppler_bin = find_poppler_bin()
    tessdata_dir = find_tessdata_dir()
    langs = []
    if tessdata_dir:
        langs = sorted(p.stem for p in tessdata_dir.glob("*.traineddata"))
    playwright_browsers = []
    if PLAYWRIGHT_BROWSERS_ROOT.exists():
        playwright_browsers = sorted(p.name for p in PLAYWRIGHT_BROWSERS_ROOT.iterdir() if p.is_dir())
    return {
        "project_root": PROJECT_ROOT,
        "python": sys.executable,
        "modules": {name: _module_ok(name) for name in REQUIRED_MODULES},
        "tesseract": tesseract,
        "tesseract_version": _command_output([str(tesseract), "--version"])
        if tesseract
        else None,
        "tessdata_dir": tessdata_dir,
        "ocr_languages": langs,
        "poppler_bin": poppler_bin,
        "pdfinfo_version": _command_output([str(poppler_bin / "pdfinfo.exe"), "-v"])
        if poppler_bin
        else None,
        "playwright_browsers_dir": PLAYWRIGHT_BROWSERS_ROOT if PLAYWRIGHT_BROWSERS_ROOT.exists() else None,
        "playwright_browsers": playwright_browsers,
    }


def main() -> int:
    console = Console()
    status = collect_status()

    console.print("[bold]Realtify environment check[/bold]")
    console.print(f"Project: {status['project_root']}")
    console.print(f"Python:  {status['python']}")

    table = Table(title="Python modules")
    table.add_column("Module")
    table.add_column("Status")
    for name, ok in status["modules"].items():
        table.add_row(name, "OK" if ok else "MISSING")
    console.print(table)

    console.print(f"Tesseract: {status['tesseract'] or 'MISSING'}")
    console.print(f"Tesseract version: {status['tesseract_version'] or 'n/a'}")
    console.print(f"Tessdata: {status['tessdata_dir'] or 'MISSING'}")
    console.print(f"OCR languages: {', '.join(status['ocr_languages']) or 'none'}")
    console.print(f"Poppler bin: {status['poppler_bin'] or 'MISSING'}")
    console.print(f"Poppler version: {status['pdfinfo_version'] or 'n/a'}")
    console.print(f"Playwright browsers: {status['playwright_browsers_dir'] or 'default user cache'}")
    console.print(f"Bundled browser dirs: {', '.join(status['playwright_browsers']) or 'none'}")

    required_langs = {"eng", "ukr", "rus", "osd"}
    ok = all(status["modules"].values())
    ok = ok and status["tesseract"] is not None
    ok = ok and status["poppler_bin"] is not None
    ok = ok and required_langs.issubset(set(status["ocr_languages"]))
    console.print("[green]Environment ready[/green]" if ok else "[red]Environment incomplete[/red]")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
