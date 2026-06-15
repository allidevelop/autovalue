from __future__ import annotations

import re
from dataclasses import dataclass

from bs4 import BeautifulSoup

from realtify.extractors.base import ListingExtractor
from realtify.models import Comparable, PropertyType, TransactionType
from realtify.screenshot_sources import PageSnapshot
from realtify.source_config import SourceDefinition


_SPACE_RE = re.compile(r"\s+")
_AREA_RE = re.compile(
    r"(?P<value>\d{1,4}(?:[.,]\d{1,2})?)\s*(?:м²|м2|м\.?\s*кв\.?|кв\.?\s*м|sqm)",
    re.IGNORECASE,
)
_ROOMS_RE = re.compile(r"(?P<value>\d{1,2})\s*(?:кімнат|кімн\.?|кім\.?|комнат|комн\.?)", re.IGNORECASE)
_MONEY_RE = re.compile(
    r"(?:(?P<prefix>\$|€|грн|uah|usd|дол\.?|доларів|євро)\s*)?"
    r"(?P<amount>\d[\d\s.,]{2,})"
    r"(?:\s*(?P<suffix>\$|€|грн|uah|usd|дол\.?|доларів|євро))?",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class ParsedMoney:
    amount: float
    currency: str | None


class GenericListingExtractor(ListingExtractor):
    def __init__(self, *, source_name: str) -> None:
        self.source_name = source_name

    def extract(
        self,
        snapshot: PageSnapshot,
        source: SourceDefinition,
        *,
        property_type: PropertyType,
        transaction_type: TransactionType,
    ) -> Comparable:
        soup = BeautifulSoup(snapshot.html, "lxml")
        text = _clean_text(snapshot.text)
        title = _first_nonempty(
            _meta(soup, "og:title"),
            _meta(soup, "twitter:title"),
            _h1(soup),
            snapshot.title,
        )
        description = _first_nonempty(
            _meta(soup, "og:description"),
            _meta(soup, "description", attr="name"),
            "",
        )
        parse_text = _clean_text("\n".join([title or "", description or "", text[:12000]]))
        money = _find_price(parse_text)
        area_m2 = _find_area(parse_text)
        rooms = _find_rooms(parse_text)
        address = _find_address(parse_text)

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

        return Comparable(
            source_url=snapshot.final_url,
            source_key=self.source_name,
            source_name=source.display_name,
            property_type=property_type,
            transaction_type=transaction_type,
            title=title,
            address=address,
            area_m2=area_m2,
            price=price,
            currency=currency,
            price_usd=price_usd,
            price_per_m2_usd=(round(price_usd / area_m2, 2) if price_usd and area_m2 else None),
            rooms=rooms,
            screenshot_path=snapshot.screenshot_path,
            warnings=warnings,
        )


def _clean_text(value: str) -> str:
    return _SPACE_RE.sub(" ", value.replace("\xa0", " ")).strip()


def _meta(soup: BeautifulSoup, key: str, *, attr: str = "property") -> str | None:
    tag = soup.find("meta", attrs={attr: key})
    if not tag and attr == "property":
        tag = soup.find("meta", attrs={"name": key})
    if not tag:
        return None
    content = tag.get("content")
    return _clean_text(content) if content else None


def _h1(soup: BeautifulSoup) -> str | None:
    tag = soup.find("h1")
    if not tag:
        return None
    return _clean_text(tag.get_text(" "))


def _first_nonempty(*values: str | None) -> str | None:
    for value in values:
        if value and value.strip():
            return _clean_text(value)
    return None


def _find_area(text: str) -> float | None:
    match = _AREA_RE.search(text)
    if not match:
        return None
    return _parse_decimal(match.group("value"))


def _find_rooms(text: str) -> int | None:
    match = _ROOMS_RE.search(text)
    if not match:
        return None
    return int(match.group("value"))


def _find_price(text: str) -> ParsedMoney | None:
    matches: list[tuple[int, ParsedMoney]] = []
    for match in _MONEY_RE.finditer(text):
        currency = _normalize_currency(match.group("prefix") or match.group("suffix"))
        if currency is None:
            continue
        amount = _parse_money_amount(match.group("amount"))
        if amount is None or amount < 100:
            continue
        start = max(0, match.start() - 80)
        end = min(len(text), match.end() + 80)
        context = text[start:end].lower()
        score = 0
        if any(word in context for word in ["ціна", "цена", "вартість", "стоимость", "price"]):
            score += 20
        if currency == "USD":
            score += 10
        if amount >= 1000:
            score += 5
        matches.append((score, ParsedMoney(amount=amount, currency=currency)))
    if not matches:
        return None
    matches.sort(key=lambda item: item[0], reverse=True)
    return matches[0][1]


def _find_address(text: str) -> str | None:
    separators = re.split(r"(?<=[.!?])\s+|\n", text)
    keywords = [
        "вул",
        "улиц",
        "просп",
        "набереж",
        "шосе",
        "площа",
        "м. ",
        "місто",
        "район",
        "жк",
    ]
    for part in separators:
        candidate = _clean_text(part)
        lowered = candidate.lower()
        if 12 <= len(candidate) <= 180 and any(keyword in lowered for keyword in keywords):
            return candidate
    return None


def _parse_decimal(value: str) -> float | None:
    normalized = value.replace(" ", "").replace(",", ".")
    try:
        return float(normalized)
    except ValueError:
        return None


def _parse_money_amount(value: str) -> float | None:
    normalized = value.replace(" ", "")
    if "," in normalized and "." in normalized:
        normalized = normalized.replace(".", "").replace(",", ".")
    elif "," in normalized:
        normalized = normalized.replace(",", ".")
    try:
        return float(normalized)
    except ValueError:
        return None


def _normalize_currency(value: str | None) -> str | None:
    if not value:
        return None
    lowered = value.lower()
    if lowered in {"$", "usd", "дол.", "доларів"}:
        return "USD"
    if lowered in {"€", "євро"}:
        return "EUR"
    if lowered in {"грн", "uah"}:
        return "UAH"
    return None
