from __future__ import annotations

import argparse
import hashlib
import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote, urlparse

from rich.console import Console

from realtify.extractors import extractor_for_source
from realtify.images import create_report_image
from realtify.models import Comparable, PropertyType, TransactionType
from realtify.paths import ensure_output_dir
from realtify.progress import ProgressCallback, emit_progress
from realtify.screenshot_sources import capture_page_snapshot
from realtify.source_config import SourcesConfig, load_sources_config


@dataclass(frozen=True)
class CollectionResult:
    output_dir: Path
    candidates: list[Comparable]
    errors: list[dict[str, str]]


def read_links(path: Path) -> list[str]:
    links: list[str] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip().lstrip("\ufeff")
        if not stripped or stripped.startswith("#"):
            continue
        links.append(stripped)
    return links


def collect_from_links(
    links: list[str],
    *,
    output_dir: Path,
    sources_config: SourcesConfig,
    property_type: PropertyType,
    transaction_type: TransactionType = "sale",
    progress: ProgressCallback | None = None,
) -> CollectionResult:
    screenshots_dir = output_dir / "screenshots"
    report_images_dir = output_dir / "report_images"
    candidates: list[Comparable] = []
    errors: list[dict[str, str]] = []

    total = len(links)
    emit_progress(progress, f"Збір оголошень: отримано {total} посилань для обробки.")
    for index, url in enumerate(links, start=1):
        source_key, source = sources_config.detect_source(url)
        screenshot_path = screenshots_dir / f"{index:02d}_{source_key}_{_url_slug(url)}.png"
        emit_progress(progress, f"[{index}/{total}] Відкриваю оголошення {source_key}: {url}")
        try:
            snapshot = capture_page_snapshot(
                url,
                screenshot_path,
                timeout_ms=sources_config.defaults.request_timeout_ms,
                viewport_width=sources_config.defaults.screenshot.viewport_width,
                viewport_height=sources_config.defaults.screenshot.viewport_height,
            )
            emit_progress(progress, f"[{index}/{total}] Скріншот збережено: {screenshot_path.name}")
            if _is_unusable_snapshot(snapshot.status_code, snapshot.title, snapshot.text):
                errors.append(
                    {
                        "url": url,
                        "source": source_key,
                        "error": f"unusable_page status={snapshot.status_code}",
                        "screenshot_path": str(screenshot_path),
                    }
                )
                emit_progress(progress, f"[{index}/{total}] Сторінка відхилена: status={snapshot.status_code}")
                continue
            extractor = extractor_for_source(source_key)
            candidate = extractor.extract(
                snapshot,
                source,
                property_type=property_type,
                transaction_type=transaction_type,
            )
            if sources_config.defaults.report_image.enabled:
                report_image_path = report_images_dir / f"{index:02d}_{source_key}_{_url_slug(url)}.jpg"
                create_report_image(
                    screenshot_path,
                    report_image_path,
                    max_width_px=sources_config.defaults.report_image.max_width_px,
                    jpeg_quality=sources_config.defaults.report_image.jpeg_quality,
                )
                candidate.report_image_path = report_image_path
                emit_progress(progress, f"[{index}/{total}] Зображення для Word підготовлено: {report_image_path.name}")
            candidates.append(candidate)
            price = f", ціна ${candidate.price_usd:,.0f}".replace(",", " ") if candidate.price_usd else ""
            area = f", площа {candidate.area_m2:g} м2" if candidate.area_m2 else ""
            emit_progress(progress, f"[{index}/{total}] Кандидат зібраний: {candidate.address or candidate.title or source_key}{area}{price}")
        except Exception as exc:
            errors.append({"url": url, "source": source_key, "error": str(exc)})
            emit_progress(progress, f"[{index}/{total}] Помилка збору: {exc}")

    emit_progress(progress, f"Збір оголошень завершено: кандидатів {len(candidates)}, помилок {len(errors)}.")
    return CollectionResult(output_dir=output_dir, candidates=candidates, errors=errors)


def save_collection_result(
    result: CollectionResult,
    *,
    links: list[str],
    candidates_filename: str = "candidates.json",
    report_filename: str = "report.md",
) -> None:
    result.output_dir.mkdir(parents=True, exist_ok=True)
    payload = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "links": links,
        "candidates": [candidate.model_dump(mode="json") for candidate in result.candidates],
        "errors": result.errors,
    }
    (result.output_dir / candidates_filename).write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (result.output_dir / report_filename).write_text(_build_report(result), encoding="utf-8")


def _build_report(result: CollectionResult) -> str:
    lines = [
        "# Comparable Collection Report",
        "",
        f"Candidates collected: {len(result.candidates)}",
        f"Errors: {len(result.errors)}",
        "",
        "## Candidates",
        "",
    ]
    for idx, candidate in enumerate(result.candidates, start=1):
        lines.extend(
            [
                f"### {idx}. {candidate.title or candidate.source_name or 'Listing'}",
                "",
                f"- Source: {candidate.source_name or 'unknown'}",
                f"- Source key: {candidate.source_key or 'unknown'}",
                f"- URL: {candidate.source_url}",
                f"- Address: {candidate.address or 'not found'}",
                f"- Area: {candidate.area_m2 if candidate.area_m2 is not None else 'not found'}",
                f"- Price: {candidate.price if candidate.price is not None else 'not found'} {candidate.currency or ''}".rstrip(),
                f"- Price USD: {candidate.price_usd if candidate.price_usd is not None else 'not found'}",
                f"- Screenshot: {candidate.screenshot_path or 'not saved'}",
                f"- Report image: {candidate.report_image_path or 'not saved'}",
                f"- Warnings: {', '.join(candidate.warnings) if candidate.warnings else 'none'}",
                "",
            ]
        )
    if result.errors:
        lines.extend(["## Errors", ""])
        for error in result.errors:
            extra = f" screenshot={error['screenshot_path']}" if error.get("screenshot_path") else ""
            lines.append(f"- `{error['source']}` {error['url']}: {error['error']}{extra}")
    return "\n".join(lines).strip() + "\n"


def _is_unusable_snapshot(status_code: int | None, title: str, text: str) -> bool:
    if status_code is not None and status_code >= 400:
        return True
    probe = f"{title}\n{text[:500]}".lower()
    blocked_markers = [
        "403 forbidden",
        "access denied",
        "captcha",
        "перевірте, що ви не робот",
        "too many requests",
        "сторінка видалена",
        "сторінку було видалено",
        "сторінка не знайдена",
        "page not found",
    ]
    return any(marker in probe for marker in blocked_markers)


def _url_slug(url: str) -> str:
    parsed = urlparse(url)
    decoded_path = unquote(parsed.path).strip("/")
    raw = decoded_path.split("/")[-1] or parsed.netloc or "listing"
    slug = re.sub(r"[^A-Za-z0-9А-Яа-яІіЇїЄєҐґ_-]+", "-", raw).strip("-")
    digest = hashlib.sha1(url.encode("utf-8")).hexdigest()[:8]
    return f"{slug[:50] or 'listing'}-{digest}"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Collect comparable listings from direct links.")
    parser.add_argument("--links", type=Path, required=True, help="UTF-8 text file with one URL per line.")
    parser.add_argument("--out", type=Path, default=None, help="Output directory. Defaults to outputs/<timestamp>.")
    parser.add_argument("--sources", type=Path, default=None, help="Path to sources.yaml.")
    parser.add_argument("--property-type", default="apartment", choices=list(PropertyType.__args__))
    parser.add_argument("--transaction-type", default="sale", choices=list(TransactionType.__args__))
    parser.add_argument("--required-count", type=int, default=None)
    args = parser.parse_args(argv)

    console = Console()
    links = read_links(args.links)
    sources_config = load_sources_config(args.sources)
    required_count = args.required_count or sources_config.defaults.required_count
    output_dir = args.out or ensure_output_dir(datetime.now().strftime("%Y%m%d_%H%M%S_links"))

    if len(links) < required_count:
        console.print(
            f"[yellow]Warning:[/yellow] got {len(links)} link(s), required count is {required_count}."
        )

    result = collect_from_links(
        links,
        output_dir=output_dir,
        sources_config=sources_config,
        property_type=args.property_type,
        transaction_type=args.transaction_type,
    )
    save_collection_result(result, links=links)

    console.print(f"[green]Collected {len(result.candidates)} candidate(s)[/green]")
    if result.errors:
        console.print(f"[red]Errors: {len(result.errors)}[/red]")
    console.print(f"Output: {result.output_dir}")
    return 0 if not result.errors else 1


if __name__ == "__main__":
    raise SystemExit(main())
