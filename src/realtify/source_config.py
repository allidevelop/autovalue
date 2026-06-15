from __future__ import annotations

from pathlib import Path
from typing import Literal
from urllib.parse import unquote, urlparse

import yaml
from pydantic import BaseModel, Field

from realtify.paths import RESOURCE_ROOT


SourceMode = Literal["direct_link", "search_url", "google_discovery"]


class ScreenshotDefaults(BaseModel):
    full_page: bool = True
    viewport_width: int = 1440
    viewport_height: int = 1200


class ReportImageDefaults(BaseModel):
    enabled: bool = True
    format: Literal["jpg", "png"] = "jpg"
    max_width_px: int = 1600
    jpeg_quality: int = 82


class SourcesDefaults(BaseModel):
    required_count: int = 5
    request_timeout_ms: int = 45000
    screenshot: ScreenshotDefaults = Field(default_factory=ScreenshotDefaults)
    report_image: ReportImageDefaults = Field(default_factory=ReportImageDefaults)


class SourceDefinition(BaseModel):
    enabled: bool = True
    priority: int = 100
    display_name: str
    domains: list[str] = Field(default_factory=list)
    modes: list[SourceMode] = Field(default_factory=lambda: ["direct_link"])
    listing_url_patterns: list[str] = Field(default_factory=list)
    search_url_templates: dict[str, list[str]] = Field(default_factory=dict)
    notes: str | None = None

    def matches_url(self, url: str) -> bool:
        parsed = urlparse(url)
        host = (parsed.netloc or "").lower()
        if host.startswith("www."):
            host_no_www = host[4:]
        else:
            host_no_www = host
        domains = {d.lower().removeprefix("www.") for d in self.domains}
        if not any(host_no_www == domain or host_no_www.endswith(f".{domain}") for domain in domains):
            return False
        if not self.listing_url_patterns:
            return True
        decoded_path = unquote(parsed.path)
        return any(
            pattern in parsed.path or pattern in decoded_path
            for pattern in self.listing_url_patterns
        )


class SourcesConfig(BaseModel):
    defaults: SourcesDefaults = Field(default_factory=SourcesDefaults)
    sources: dict[str, SourceDefinition]

    def enabled_sources(self) -> list[tuple[str, SourceDefinition]]:
        return sorted(
            [(name, source) for name, source in self.sources.items() if source.enabled],
            key=lambda item: item[1].priority,
        )

    def detect_source(self, url: str) -> tuple[str, SourceDefinition]:
        for name, source in self.enabled_sources():
            if source.matches_url(url):
                return name, source
        fallback = self.sources.get("developer_or_agency")
        if fallback and fallback.enabled:
            return "developer_or_agency", fallback
        return "unknown", SourceDefinition(
            enabled=True,
            priority=999,
            display_name="Unknown source",
            domains=[],
            modes=["direct_link"],
            listing_url_patterns=[],
        )


def load_sources_config(path: Path | None = None) -> SourcesConfig:
    config_path = path or RESOURCE_ROOT / "config" / "sources.yaml"
    with config_path.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return SourcesConfig.model_validate(data)
