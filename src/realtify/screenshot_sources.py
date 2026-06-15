from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from playwright.sync_api import sync_playwright


@dataclass(frozen=True)
class PageSnapshot:
    requested_url: str
    final_url: str
    status_code: int | None
    title: str
    html: str
    text: str
    screenshot_path: Path


def capture_full_page_screenshot(url: str, output_path: Path, *, timeout_ms: int = 45000) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1440, "height": 1200})
        page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
        page.wait_for_load_state("networkidle", timeout=timeout_ms)
        _try_close_common_banners(page)
        page.screenshot(path=str(output_path), full_page=True)
        browser.close()
    return output_path


def capture_page_snapshot(
    url: str,
    output_path: Path,
    *,
    timeout_ms: int = 45000,
    viewport_width: int = 1440,
    viewport_height: int = 1200,
) -> PageSnapshot:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": viewport_width, "height": viewport_height},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/126.0.0.0 Safari/537.36"
            ),
            locale="uk-UA",
            extra_http_headers={
                "Accept-Language": "uk-UA,uk;q=0.9,ru;q=0.8,en-US;q=0.7,en;q=0.6"
            },
        )
        page = context.new_page()
        response = page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
        try:
            page.wait_for_load_state("networkidle", timeout=timeout_ms)
        except Exception:
            pass
        _try_close_common_banners(page)
        page.screenshot(path=str(output_path), full_page=True)
        snapshot = PageSnapshot(
            requested_url=url,
            final_url=page.url,
            status_code=response.status if response else None,
            title=page.title(),
            html=page.content(),
            text=page.locator("body").inner_text(timeout=5000),
            screenshot_path=output_path,
        )
        browser.close()
        return snapshot


def _try_close_common_banners(page) -> None:
    labels = [
        "Accept",
        "I agree",
        "OK",
        "Прийняти",
        "Погоджуюсь",
        "Згоден",
        "Закрити",
        "Close",
    ]
    for label in labels:
        try:
            page.get_by_text(label, exact=False).first.click(timeout=1000)
            return
        except Exception:
            continue
