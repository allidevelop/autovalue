from __future__ import annotations

import json
import re
from collections.abc import Iterable
from typing import Any

from bs4 import BeautifulSoup

from realtify.extractors.base import ListingExtractor
from realtify.extractors.generic import (
    ParsedMoney,
    _clean_text,
    _find_price,
    _first_nonempty,
    _h1,
    _meta,
    _parse_decimal,
    _parse_money_amount,
)
from realtify.models import Comparable, PropertyType, TransactionType
from realtify.screenshot_sources import PageSnapshot
from realtify.source_config import SourceDefinition


class DimriaListingExtractor(ListingExtractor):
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
        product = _find_jsonld_product(soup)
        offer = _offer(product)

        title = _first_nonempty(
            _h1(soup),
            _string(product.get("name") if product else None),
            _meta(soup, "og:title"),
            snapshot.title,
        )
        money = _extract_price(offer, soup, "\n".join(lines[:120]))
        area_m2 = _extract_area(title or "", text)
        rooms = _extract_rooms(title or "", text)
        address = _extract_address(lines, product)
        city, district = _split_city_district(address)
        complex_name = _extract_complex_name(lines, title)
        delivery_date = _extract_delivery_date(text)
        building_class = _extract_building_class(text)
        floor = _extract_floor(text)
        price_per_m2 = _extract_price_per_m2(text)

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
            source_key="dimria",
            source_name=source.display_name,
            property_type=property_type,
            transaction_type=transaction_type,
            title=title,
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
            building_class=building_class,
            delivery_date=delivery_date,
            screenshot_path=snapshot.screenshot_path,
            warnings=warnings,
        )


def _lines(text: str) -> list[str]:
    return [_clean_text(line) for line in text.splitlines() if _clean_text(line)]


def _find_jsonld_product(soup: BeautifulSoup) -> dict[str, Any] | None:
    for obj in _jsonld_objects(soup):
        if obj.get("@type") == "Product":
            return obj
    return None


def _jsonld_objects(soup: BeautifulSoup) -> Iterable[dict[str, Any]]:
    for script in soup.find_all("script", type="application/ld+json"):
        raw = script.get_text(strip=True)
        if not raw:
            continue
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            continue
        yield from _walk_dicts(parsed)


def _walk_dicts(value: Any) -> Iterable[dict[str, Any]]:
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from _walk_dicts(child)
    elif isinstance(value, list):
        for item in value:
            yield from _walk_dicts(item)


def _offer(product: dict[str, Any] | None) -> dict[str, Any] | None:
    if not product:
        return None
    offer = product.get("offers")
    return offer if isinstance(offer, dict) else None


def _extract_price(offer: dict[str, Any] | None, soup: BeautifulSoup, text_head: str) -> ParsedMoney | None:
    if offer:
        amount = _parse_money_amount(str(offer.get("price") or ""))
        currency = _normalize_currency(str(offer.get("priceCurrency") or ""))
        if amount and currency:
            return ParsedMoney(amount=amount, currency=currency)
    selector = soup.select_one(".price")
    selector_text = _clean_text(selector.get_text(" ")) if selector else ""
    return _find_price(selector_text or text_head)


def _normalize_currency(value: str | None) -> str | None:
    if not value:
        return None
    normalized = value.strip().upper()
    if normalized in {"USD", "$", "ДОЛ", "ДОЛАРІВ"}:
        return "USD"
    if normalized in {"UAH", "ГРН"}:
        return "UAH"
    if normalized in {"EUR", "€", "ЄВРО"}:
        return "EUR"
    return None


def _extract_area(title: str, text: str) -> float | None:
    for pattern in [
        r"Загальна\s+площа\s*(?P<value>\d{1,4}(?:[.,]\d+)?)\s*м²",
        r"(?P<value>\d{1,4}(?:[.,]\d+)?)\s*кв\.\s*м",
        r"Площа\s*(?P<value>\d{1,4}(?:[.,]\d+)?)\s*м²",
    ]:
        match = re.search(pattern, f"{title}\n{text}", re.I)
        if match:
            return _parse_decimal(match.group("value"))
    return None


def _extract_rooms(title: str, text: str) -> int | None:
    match = re.search(r"Продаж\s+(?P<value>\d{1,2})к", title, re.I)
    if match:
        return int(match.group("value"))
    match = re.search(r"(?P<value>\d{1,2})\s*кімнат", text, re.I)
    if match:
        return int(match.group("value"))
    return None


def _extract_address(lines: list[str], product: dict[str, Any] | None) -> str | None:
    for line in lines[:120]:
        if "·" in line and any(city in line for city in ["Львів", "Київ", "Одеса", "Дніпро", "Харків"]):
            if "$" not in line and "грн" not in line.casefold():
                return _clean_text(line.replace(" · ", ", ").replace("·", ", "))
    offer = _offer(product)
    if offer and offer.get("areaServed"):
        return _clean_text(str(offer["areaServed"]))
    return None


def _split_city_district(address: str | None) -> tuple[str | None, str | None]:
    if not address:
        return None, None
    parts = [_clean_text(part) for part in address.split(",") if _clean_text(part)]
    city = next((part for part in reversed(parts) if part in {"Львів", "Київ", "Одеса", "Дніпро", "Харків"}), None)
    district = next(
        (
            part
            for part in parts
            if re.search(r"(^|\s)район\b|р-н", part.casefold())
        ),
        None,
    )
    return city, district


def _extract_complex_name(lines: list[str], title: str | None) -> str | None:
    if title:
        try:
            idx = lines.index(title)
        except ValueError:
            idx = -1
        if idx >= 0 and idx + 1 < len(lines) and re.match(r"Ж[КБ]\s+", lines[idx + 1]):
            return lines[idx + 1]
    for line in lines[:120]:
        if re.match(r"Ж[КБ]\s+[A-Za-zА-Яа-яІіЇїЄєҐґ0-9'’ -]{2,80}$", line):
            return line
    return None


def _extract_delivery_date(text: str) -> str | None:
    match = re.search(r"(Здача\s+в\s+[^\n\r.]+|Побудовано\s+\d{4})", text, re.I)
    if not match:
        return None
    return _clean_text(match.group(1))


def _extract_building_class(text: str) -> str | None:
    match = re.search(r"Клас\s+(?P<value>економ|комфорт|бізнес|преміум)", text, re.I)
    if not match:
        return None
    return match.group("value").casefold()


def _extract_floor(text: str) -> str | None:
    match = re.search(r"(?P<floor>\d{1,3})\s+поверх\s+з\s+(?P<total>\d{1,3})", text, re.I)
    if not match:
        return None
    return f"{match.group('floor')} з {match.group('total')}"


def _extract_price_per_m2(text: str) -> float | None:
    match = re.search(r"(?P<value>\d[\d\s.,]*)\s*\$\s+за\s+м²", text, re.I)
    if not match:
        return None
    return _parse_decimal(match.group("value"))


def _string(value: Any) -> str | None:
    return value if isinstance(value, str) else None
