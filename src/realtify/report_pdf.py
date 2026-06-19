"""HTML → PDF через headless-Chromium (Playwright page.pdf). Поля/розмір сторінки
беруться з CSS @page (prefer_css_page_size), тобто з тієї ж style-spec, що й docx.
Жодних нових залежностей: playwright уже у проєкті (див. screenshot_sources.py)."""
from __future__ import annotations

from pathlib import Path
from typing import Any

from realtify import paths  # noqa: F401 — side-effect: виставляє PLAYWRIGHT_BROWSERS_PATH
from realtify import report_html, report_styles as styles


def render_html_to_pdf(html: str, out_path: Path) -> Path:
    from playwright.sync_api import sync_playwright

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            page = browser.new_page()
            page.set_content(html, wait_until="networkidle")
            page.pdf(path=str(out_path), print_background=True, prefer_css_page_size=True)
        finally:
            browser.close()
    return out_path


def render_document_pdf(document: dict, out_path: Path, spec: dict[str, Any] | None = None, mode: str = "clean") -> Path:
    spec = spec or styles.load_style_spec()
    html = report_html.render_document_html(document, spec, mode)
    return render_html_to_pdf(html, out_path)
