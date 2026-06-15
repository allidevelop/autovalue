from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, Field, HttpUrl


PropertyType = Literal[
    "parking",
    "apartment",
    "commercial",
    "office",
    "retail",
    "warehouse",
    "house",
    "land",
]

TransactionType = Literal["sale", "rent"]


class Comparable(BaseModel):
    source_url: HttpUrl
    source_key: str | None = None
    source_name: str | None = None
    property_type: PropertyType
    transaction_type: TransactionType = "sale"
    title: str | None = None
    address: str | None = None
    city: str | None = None
    district: str | None = None
    complex_name: str | None = None
    area_m2: float | None = None
    price: float | None = None
    currency: str | None = None
    price_usd: float | None = None
    price_per_m2_usd: float | None = None
    floor_or_level: str | None = None
    rooms: int | None = None
    condition: str | None = None
    building_class: str | None = None
    purpose: str | None = None
    delivery_date: str | None = None
    location_quality: str | None = None
    screenshot_path: Path | None = None
    report_image_path: Path | None = None
    collected_at: datetime = Field(default_factory=datetime.now)
    warnings: list[str] = Field(default_factory=list)


class BargainingAdjustment(BaseModel):
    enabled: bool = False
    exposure_months: float
    annual_discount_rate_pct: float
    annual_market_growth_pct: float
    round_to_pct: float = 1

    def raw_discount_pct(self) -> float:
        return self.exposure_months * (
            self.annual_discount_rate_pct - self.annual_market_growth_pct
        ) / 12

    def excel_adjustment_pct(self) -> float:
        raw = self.raw_discount_pct()
        if self.round_to_pct <= 0:
            return -raw
        rounded = round(raw / self.round_to_pct) * self.round_to_pct
        return -rounded
