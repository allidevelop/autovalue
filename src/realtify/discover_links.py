from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import parse_qsl, urlencode, urljoin, urlparse, urlunparse

import yaml
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright
from pydantic import BaseModel, HttpUrl
from rich.console import Console

from realtify.models import PropertyType, TransactionType
from realtify.paths import PROJECT_ROOT, ensure_output_dir
from realtify.progress import ProgressCallback, emit_progress
from realtify.source_config import SourcesConfig, load_sources_config


class DiscoveryError(RuntimeError):
    pass


class DiscoveredLink(BaseModel):
    url: HttpUrl
    source_key: str
    source_name: str | None = None
    source_page_url: str
    rank: int


class DiscoverySourcePage(BaseModel):
    source_key: str
    url: str
    final_url: str | None = None
    status_code: int | None = None
    title: str | None = None
    listing_links_found: int = 0
    error: str | None = None


@dataclass(frozen=True)
class DiscoveryResult:
    output_dir: Path
    links: list[DiscoveredLink]
    source_pages: list[DiscoverySourcePage]
    rejected: list[dict[str, str]]
    warnings: list[str]


CITY_SLUGS = {
    "київ": "kiev",
    "киев": "kiev",
    "kyiv": "kiev",
    "kiev": "kiev",
    "львів": "lvov",
    "львов": "lvov",
    "lviv": "lvov",
    "lvov": "lvov",
    "одеса": "odessa",
    "одесса": "odessa",
    "odesa": "odessa",
    "odessa": "odessa",
    "дніпро": "dnepr",
    "днепр": "dnepr",
    "dnipro": "dnepr",
    "dnepr": "dnepr",
    "харків": "kharkov",
    "харьков": "kharkov",
    "kharkiv": "kharkov",
    "kharkov": "kharkov",
    "ужгород": "uzhgorod",
    "uzhhorod": "uzhgorod",
    "uzhgorod": "uzhgorod",
    "івано-франківськ": "ivano-frankovsk",
    "ивано-франковск": "ivano-frankovsk",
    "ivano-frankivsk": "ivano-frankovsk",
    "тернопіль": "ternopol",
    "тернополь": "ternopol",
    "ternopil": "ternopol",
}


def discover_links_for_task(
    *,
    task: dict[str, Any],
    output_dir: Path,
    sources_config: SourcesConfig,
    required_count: int | None = None,
    max_links: int | None = None,
    pages_per_source: int | None = None,
    progress: ProgressCallback | None = None,
) -> DiscoveryResult:
    target = _section(task, "target")
    collection = _section(task, "collection")

    property_type = _typed_property_type(target.get("property_type") or "apartment")
    transaction_type = _typed_transaction_type(target.get("transaction_type") or "sale")
    only_newbuilds = _optional_bool(collection.get("only_newbuilds"), default=True)
    required = required_count or _optional_int(collection.get("required_count")) or sources_config.defaults.required_count
    limit = max_links or _optional_int(collection.get("max_discovered_links")) or max(required * 4, required)
    page_count = pages_per_source or _optional_int(collection.get("discovery_pages")) or 1

    custom_pages = _custom_source_pages(collection.get("search_urls"), sources_config)
    if _optional_bool(collection.get("search_only_custom"), default=False) and custom_pages:
        # Примусовий пошук у конкретному ЖК/будинку: лише задані каталог-URL, без city-wide.
        planned_pages, warnings = list(custom_pages), []
    else:
        planned_pages, warnings = build_source_pages(
            target=target,
            sources_config=sources_config,
            property_type=property_type,
            transaction_type=transaction_type,
            only_newbuilds=only_newbuilds,
            pages_per_source=page_count,
        )
        planned_pages.extend(custom_pages)
    if not planned_pages:
        raise DiscoveryError("No source search pages could be built. Check target.city, target.property_type, and config/sources.yaml.")

    emit_progress(progress, f"Пошук посилань: заплановано {len(planned_pages)} сторінок каталогів, ліміт {limit} посилань.")
    result = _collect_listing_links(
        planned_pages=planned_pages,
        sources_config=sources_config,
        output_dir=output_dir,
        max_links=limit,
        warnings=warnings,
        progress=progress,
    )
    if len(result.links) < required:
        result.warnings.append(f"discovered_only_{len(result.links)}_links_required_{required}")
    return result


def build_source_pages(
    *,
    target: dict[str, Any],
    sources_config: SourcesConfig,
    property_type: PropertyType,
    transaction_type: TransactionType,
    only_newbuilds: bool,
    pages_per_source: int,
) -> tuple[list[tuple[str, str]], list[str]]:
    warnings: list[str] = []
    city_slug = _city_slug(target)
    rooms = _optional_int(target.get("rooms"))
    planned: list[tuple[str, str]] = []

    for source_key, source in sources_config.enabled_sources():
        if "search_url" not in source.modes:
            continue
        template_keys = _template_keys(property_type, transaction_type, only_newbuilds)
        templates: list[str] = []
        for key in template_keys:
            values = source.search_url_templates.get(key, [])
            templates.extend(values)
            if templates:
                break
        if only_newbuilds and not templates and source.search_url_templates.get(f"{property_type}_{transaction_type}"):
            warnings.append(f"{source_key}: skipped because no strict newbuild search URL is configured")
        for template in templates:
            url = _format_template(template, city_slug=city_slug, rooms=rooms)
            for page_number in range(1, max(1, pages_per_source) + 1):
                planned.append((source_key, _with_page_number(url, page_number)))
    return planned, warnings


def save_discovery_result(result: DiscoveryResult) -> Path:
    result.output_dir.mkdir(parents=True, exist_ok=True)
    links_path = result.output_dir / "discovered_links.txt"
    links_path.write_text(
        "\n".join(str(link.url) for link in result.links) + ("\n" if result.links else ""),
        encoding="utf-8",
    )
    payload = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "links_file": str(links_path),
        "links": [link.model_dump(mode="json") for link in result.links],
        "source_pages": [page.model_dump(mode="json") for page in result.source_pages],
        "rejected": result.rejected,
        "warnings": result.warnings,
    }
    (result.output_dir / "discovery.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return links_path


def _collect_listing_links(
    *,
    planned_pages: list[tuple[str, str]],
    sources_config: SourcesConfig,
    output_dir: Path,
    max_links: int,
    warnings: list[str],
    progress: ProgressCallback | None = None,
) -> DiscoveryResult:
    links: list[DiscoveredLink] = []
    source_pages: list[DiscoverySourcePage] = []
    rejected: list[dict[str, str]] = []
    seen: set[str] = set()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1440, "height": 1200},
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
        for page_index, (expected_source_key, page_url) in enumerate(planned_pages, start=1):
            page_record = DiscoverySourcePage(source_key=expected_source_key, url=page_url)
            emit_progress(progress, f"[Discovery {page_index}/{len(planned_pages)}] Відкриваю каталог {expected_source_key}: {page_url}")
            try:
                response = page.goto(page_url, wait_until="domcontentloaded", timeout=45000)
                page_record.status_code = response.status if response else None
                try:
                    page.wait_for_load_state("networkidle", timeout=15000)
                except Exception:
                    pass
                _scroll_catalog_page(page)
                page_record.final_url = page.url
                page_record.title = page.title()
                found = _extract_listing_urls(page.content(), page.url, sources_config, expected_source_key)
                page_record.listing_links_found = len(found)
                emit_progress(progress, f"[Discovery {page_index}/{len(planned_pages)}] Знайдено посилань на сторінці: {len(found)}")
                for url in found:
                    normalized = _normalize_url(url)
                    if normalized in seen:
                        continue
                    source_key, source = sources_config.detect_source(normalized)
                    if source_key != expected_source_key:
                        rejected.append({"url": normalized, "reason": f"unexpected_source:{source_key}"})
                        continue
                    if not source.matches_url(normalized):
                        rejected.append({"url": normalized, "reason": "not_listing_url"})
                        continue
                    seen.add(normalized)
                    links.append(
                        DiscoveredLink(
                            url=normalized,
                            source_key=source_key,
                            source_name=source.display_name,
                            source_page_url=page.url,
                            rank=len(links) + 1,
                        )
                    )
                    if len(links) >= max_links:
                        break
            except Exception as exc:
                page_record.error = str(exc)
                emit_progress(progress, f"[Discovery {page_index}/{len(planned_pages)}] Помилка каталогу: {exc}")
            source_pages.append(page_record)
            if len(links) >= max_links:
                break
        browser.close()

    emit_progress(progress, f"Пошук посилань завершено: відібрано {len(links)} унікальних URL.")
    return DiscoveryResult(
        output_dir=output_dir,
        links=links,
        source_pages=source_pages,
        rejected=rejected,
        warnings=warnings,
    )


def _extract_listing_urls(
    html: str,
    base_url: str,
    sources_config: SourcesConfig,
    expected_source_key: str,
) -> list[str]:
    soup = BeautifulSoup(html, "lxml")
    urls: list[str] = []
    for tag in soup.find_all("a", href=True):
        href = str(tag.get("href") or "").strip()
        for candidate in _urls_from_href(href, base_url):
            source_key, source = sources_config.detect_source(candidate)
            if source_key == expected_source_key and source.matches_url(candidate):
                urls.append(candidate)
    return _dedupe(urls)


def _urls_from_href(href: str, base_url: str) -> list[str]:
    if not href:
        return []
    if href.startswith(("tel:", "mailto:", "#")):
        return []
    if href.startswith(("viber:", "tg:", "whatsapp:")):
        return re.findall(r"https?://[^\s&\"']+", href)
    absolute = urljoin(base_url, href)
    if not absolute.startswith(("http://", "https://")):
        return []
    return [absolute]


def _scroll_catalog_page(page) -> None:
    for fraction in (0.35, 0.7, 1.0):
        try:
            page.evaluate("(fraction) => window.scrollTo(0, document.body.scrollHeight * fraction)", fraction)
            page.wait_for_timeout(700)
        except Exception:
            return


def _template_keys(
    property_type: PropertyType,
    transaction_type: TransactionType,
    only_newbuilds: bool,
) -> list[str]:
    base = f"{property_type}_{transaction_type}"
    if only_newbuilds:
        return [f"{base}_newbuild"]
    return [base]


def _format_template(template: str, *, city_slug: str, rooms: int | None) -> str:
    rooms_query = f"?rooms={rooms}" if rooms else ""
    newbuild_rooms_query = f"?newhouse=1&rooms={rooms}" if rooms else "?newhouse=1"
    return template.format(
        city_slug=city_slug,
        rooms=rooms or "",
        rooms_query=rooms_query,
        newbuild_rooms_query=newbuild_rooms_query,
    )


def _with_page_number(url: str, page_number: int) -> str:
    if page_number <= 1:
        return url
    parsed = urlparse(url)
    query = dict(parse_qsl(parsed.query, keep_blank_values=True))
    query["page"] = str(page_number)
    return urlunparse(parsed._replace(query=urlencode(query)))


def _normalize_url(url: str) -> str:
    parsed = urlparse(url)
    query = [
        (key, value)
        for key, value in parse_qsl(parsed.query, keep_blank_values=True)
        if not key.lower().startswith("utm_")
    ]
    path = re.sub(r"/{2,}", "/", parsed.path)
    return urlunparse(
        parsed._replace(
            scheme=parsed.scheme.lower(),
            netloc=parsed.netloc.lower(),
            path=path,
            query=urlencode(query),
            fragment="",
        )
    )


def _dedupe(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        normalized = _normalize_url(value)
        if normalized in seen:
            continue
        seen.add(normalized)
        result.append(normalized)
    return result


def _city_slug(target: dict[str, Any]) -> str:
    city = str(target.get("city") or "")
    address = str(target.get("address") or "")
    probe = _normalize_city_probe(f"{city} {address}")
    for key, slug in CITY_SLUGS.items():
        if key in probe:
            return slug
    raise DiscoveryError(f"Unsupported or missing city for source discovery: {city or address}")


def _normalize_city_probe(value: str) -> str:
    lowered = value.casefold()
    lowered = re.sub(r"\bм\.\s*", " ", lowered)
    lowered = lowered.replace("ё", "е")
    return re.sub(r"\s+", " ", lowered)


def _custom_source_pages(value: Any, sources_config: SourcesConfig) -> list[tuple[str, str]]:
    if not value:
        return []
    if not isinstance(value, list):
        raise DiscoveryError("collection.search_urls must be a list")
    planned: list[tuple[str, str]] = []
    for item in value:
        url = str(item).strip()
        if not url:
            continue
        source_key, _source = sources_config.detect_source(url)
        if source_key == "developer_or_agency":
            continue
        planned.append((source_key, url))
    return planned


def _load_yaml(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    if not isinstance(data, dict):
        raise DiscoveryError(f"{path} must contain a YAML object")
    return data


def _section(task: dict[str, Any], name: str) -> dict[str, Any]:
    value = task.get(name) or {}
    if not isinstance(value, dict):
        raise DiscoveryError(f"task.{name} must be a YAML object")
    return value


def _resolve_path(path: Path) -> Path:
    return path if path.is_absolute() else PROJECT_ROOT / path


def _optional_int(value: Any) -> int | None:
    if value in (None, ""):
        return None
    return int(value)


def _optional_bool(value: Any, *, default: bool) -> bool:
    if value in (None, ""):
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().casefold() in {"1", "true", "yes", "y", "так"}


def _typed_property_type(value: Any) -> PropertyType:
    if value not in PropertyType.__args__:
        raise DiscoveryError(f"Unsupported property_type: {value}")
    return value


def _typed_transaction_type(value: Any) -> TransactionType:
    if value not in TransactionType.__args__:
        raise DiscoveryError(f"Unsupported transaction_type: {value}")
    return value


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Discover comparable listing URLs from configured source catalogs.")
    parser.add_argument("--task", type=Path, required=True, help="Task YAML file with target and collection settings.")
    parser.add_argument("--out", type=Path, default=None, help="Output directory. Defaults to outputs/<timestamp>_discovery.")
    parser.add_argument("--sources", type=Path, default=None, help="Path to sources.yaml.")
    parser.add_argument("--required-count", type=int, default=None)
    parser.add_argument("--max-links", type=int, default=None)
    parser.add_argument("--pages", type=int, default=None, help="Catalog pages per source.")
    args = parser.parse_args(argv)

    console = Console()
    try:
        task_file = _resolve_path(args.task)
        task = _load_yaml(task_file)
        sources_config = load_sources_config(_resolve_path(args.sources) if args.sources else None)
        output_dir = _resolve_path(args.out) if args.out else ensure_output_dir(datetime.now().strftime("%Y%m%d_%H%M%S_discovery"))
        result = discover_links_for_task(
            task=task,
            output_dir=output_dir,
            sources_config=sources_config,
            required_count=args.required_count,
            max_links=args.max_links,
            pages_per_source=args.pages,
        )
        links_path = save_discovery_result(result)
    except Exception as exc:
        console.print(f"[red]Discovery failed:[/red] {exc}")
        return 1

    console.print(f"[green]Discovered {len(result.links)} link(s)[/green]")
    if result.warnings:
        console.print(f"[yellow]Warnings: {len(result.warnings)}[/yellow]")
    console.print(f"Links: {links_path}")
    console.print(f"Output: {result.output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
