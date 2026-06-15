from __future__ import annotations

import re

from bs4 import BeautifulSoup

from realtify.extractors.base import ListingExtractor
from realtify.extractors.generic import (
    _clean_text,
    _find_price,
    _first_nonempty,
    _meta,
    _parse_decimal,
)
from realtify.models import Comparable, PropertyType, TransactionType
from realtify.screenshot_sources import PageSnapshot
from realtify.source_config import SourceDefinition


class RieltorListingExtractor(ListingExtractor):
    def extract(
        self,
        snapshot: PageSnapshot,
        source: SourceDefinition,
        *,
        property_type: PropertyType,
        transaction_type: TransactionType,
    ) -> Comparable:
        soup = BeautifulSoup(snapshot.html, "lxml")
        text = snapshot.text
        lines = _lines(text)

        title = _clean_title(
            _first_nonempty(_meta(soup, "og:title"), _selector_text(soup, "title"), snapshot.title)
        )
        description = _first_nonempty(
            _meta(soup, "og:description"),
            _meta(soup, "description", attr="name"),
            "",
        )
        money = _extract_price(soup, "\n".join(lines[:80]))
        area_m2 = _extract_area(text)
        rooms = _extract_rooms(text)
        address = _extract_address(soup, lines)
        city, district = _split_city_district(address)
        price_per_m2 = _extract_price_per_m2(soup, "\n".join(lines[:80]))
        complex_name = _extract_complex_name(lines)
        condition = _extract_labeled_value(text, "Загальний стан квартири")
        building_class = _extract_building_class(lines[:80])
        floor = _extract_floor(text)

        warnings: list[str] = []
        if money is None:
            warnings.append("price_not_found")
        if area_m2 is None:
            warnings.append("area_not_found")
        if address is None:
            warnings.append("address_not_found")

        price = money.amount if money else None
        currency = money.currency if money else None
        price_usd = price if currency == "USD" else None
        if price and currency and currency != "USD":
            warnings.append(f"price_currency_not_usd:{currency}")
        if price_per_m2 is None and price_usd and area_m2:
            price_per_m2 = round(price_usd / area_m2, 2)

        return Comparable(
            source_url=snapshot.final_url,
            source_key="rieltor",
            source_name=source.display_name,
            property_type=property_type,
            transaction_type=transaction_type,
            title=title or description,
            address=address,
            city=city,
            district=district,
            complex_name=complex_name,
            area_m2=area_m2,
            price=price,
            currency=currency,
            price_usd=price_usd,
            price_per_m2_usd=price_per_m2,
            floor_or_level=floor,
            rooms=rooms,
            condition=condition,
            building_class=building_class,
            screenshot_path=snapshot.screenshot_path,
            warnings=warnings,
        )


def _selector_text(soup: BeautifulSoup, selector: str) -> str | None:
    tag = soup.select_one(selector)
    if not tag:
        return None
    return _clean_text(tag.get_text(" "))


def _clean_title(value: str | None) -> str | None:
    if not value:
        return None
    title = re.sub(r"\s+-\s+RIELTOR\.UA\s*$", "", value, flags=re.IGNORECASE)
    return _clean_text(title)


def _lines(text: str) -> list[str]:
    return [_clean_text(line) for line in text.splitlines() if _clean_text(line)]


def _extract_price(soup: BeautifulSoup, text_head: str):
    selector_value = _first_nonempty(
        _selector_text(soup, ".offer-view-price-title"),
        _selector_text(soup, ".offer-view-price"),
    )
    money = _find_price(selector_value or "")
    if money:
        return money
    return _find_price(text_head)


def _extract_price_per_m2(soup: BeautifulSoup, text_head: str) -> float | None:
    value = _first_nonempty(
        _selector_text(soup, ".offer-view-price-subtitle"),
        _selector_text(soup, ".offer-view-price"),
        text_head,
    )
    if not value:
        return None
    match = re.search(r"(?P<value>\d[\d\s.,]*)\s*\$/\s*м²", value)
    if not match:
        return None
    return _parse_decimal(match.group("value"))


def _extract_area(text: str) -> float | None:
    label_match = re.search(r"Загальна\s+площа\s*:\s*(?P<value>\d+(?:[.,]\d+)?)\s*м²", text, re.I)
    if label_match:
        return _parse_decimal(label_match.group("value"))
    layout_match = re.search(
        r"(?P<total>\d{1,4}(?:[.,]\d+)?)\s*/\s*\d{1,4}(?:[.,]\d+)?\s*/\s*\d{1,4}(?:[.,]\d+)?\s*м²",
        text,
        re.I,
    )
    if layout_match:
        return _parse_decimal(layout_match.group("total"))
    plain_match = re.search(r"загальною\s+площею\s+(?P<value>\d+(?:[.,]\d+)?)\s*(?:кв\.\s*метрів|м²)", text, re.I)
    if plain_match:
        return _parse_decimal(plain_match.group("value"))
    return None


def _extract_rooms(text: str) -> int | None:
    label_match = re.search(r"Кількість\s+кімнат\s*:\s*(?P<value>\d{1,2})", text, re.I)
    if label_match:
        return int(label_match.group("value"))
    match = re.search(r"(?P<value>\d{1,2})\s*кімнат", text, re.I)
    if not match:
        return None
    return int(match.group("value"))


def _extract_address(soup: BeautifulSoup, lines: list[str]) -> str | None:
    for idx, line in enumerate(lines):
        if line.casefold() == "адреса" and idx + 1 < len(lines):
            return lines[idx + 1]
    selector_value = _selector_text(soup, ".offer-view-address")
    if selector_value:
        city = _nearby_line(lines, selector_value, contains_any=["Львів", "Київ", "Одеса", "Дніпро", "Харків"])
        district = _nearby_line(lines, selector_value, contains_any=["р-н", "район"])
        parts = [part for part in [city, district, selector_value] if part]
        return ", ".join(dict.fromkeys(parts))
    return None


def _nearby_line(lines: list[str], anchor: str, *, contains_any: list[str]) -> str | None:
    try:
        idx = lines.index(anchor)
    except ValueError:
        return None
    for line in lines[idx + 1 : idx + 6]:
        if any(marker in line for marker in contains_any):
            return line.replace(" ,", ",")
    return None


def _split_city_district(address: str | None) -> tuple[str | None, str | None]:
    if not address:
        return None, None
    parts = [_clean_text(part) for part in address.split(",") if _clean_text(part)]
    city = parts[0] if parts else None
    district = next((part for part in parts if "р-н" in part or "район" in part.casefold()), None)
    return city, district


def _extract_complex_name(lines: list[str]) -> str | None:
    for line in lines[:120]:
        if line.count("ЖК") > 1:
            continue
        match = re.search(r"\bЖК\s+[A-Za-zА-Яа-яІіЇїЄєҐґ0-9'’ -]{2,80}", line)
        if match:
            return _clean_text(match.group(0))
    return None


def _extract_labeled_value(text: str, label: str) -> str | None:
    match = re.search(rf"{re.escape(label)}\s*:\s*(?P<value>[^\n\r]+)", text, re.I)
    if not match:
        return None
    return _clean_text(match.group("value"))


def _extract_building_class(lines: list[str]) -> str | None:
    for line in lines:
        lowered = line.casefold()
        for value in ["економ", "комфорт", "бізнес", "преміум"]:
            if value in lowered:
                return value
    return None


def _extract_floor(text: str) -> str | None:
    match = re.search(r"поверх\s+(?P<floor>\d{1,3})\s+з\s+(?P<total>\d{1,3})", text, re.I)
    if match:
        return f"{match.group('floor')} з {match.group('total')}"
    floor = _extract_labeled_value(text, "Поверх")
    total = _extract_labeled_value(text, "Поверховість")
    if floor and total:
        return f"{floor} з {total}"
    return floor
