"""Best-effort резолв URL каталогу ЖК/будинку для примусового пошуку аналогів.

Надійний шлях — користувач вставляє посилання на каталог ЖК (поле у формі).
Цей модуль — лише авто-спроба побудувати сторінку будинку на dom.ria з адреси
об'єкта і ПЕРЕВІРИТИ її фетчем (повертаємо URL тільки якщо реально відкрився
і містить оголошення). Якщо не вдалось — None, і воркфлоу розширює пошук
з чесною позначкою «орієнтовна».

Без нейромереж: детермінований транслітерат + верифікація HTTP-запитом.
"""

from __future__ import annotations

import re
from typing import Any

import requests

_TIMEOUT = 12
_DOMRIA_BUILDING_TMPL = "https://dom.ria.com/uk/prodazha-kvartir/{city}-{street}-zdanie-{building}/"
_LISTING_RE = re.compile(r"realty-prodaja-kvartira[a-z0-9-]+\.html")
_MIN_LISTINGS = 3

# Місто → слаг dom.ria (як у CITY_SLUGS discover_links).
_CITY_SLUGS = {
    "київ": "kiev", "киев": "kiev", "kyiv": "kiev", "kiev": "kiev",
    "львів": "lvov", "львов": "lvov", "lviv": "lvov", "lvov": "lvov",
    "одеса": "odessa", "одесса": "odessa", "odesa": "odessa",
    "дніпро": "dnepr", "днепр": "dnepr",
    "харків": "kharkov", "харьков": "kharkov",
}

# Транслітерат для слага вулиці/літери будинку (кирилиця → латиниця).
_TRANSLIT = {
    "а": "a", "б": "b", "в": "v", "г": "g", "ґ": "g", "д": "d", "е": "e", "є": "ie",
    "ж": "zh", "з": "z", "и": "y", "і": "i", "ї": "i", "й": "y", "к": "k", "л": "l",
    "м": "m", "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ф": "f", "х": "h", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "shch", "ь": "",
    "ю": "iu", "я": "ia", "’": "", "'": "",
}


def resolve_complex_catalog_url(target: dict[str, Any]) -> str | None:
    """Авто-спроба знайти сторінку будинку на dom.ria за адресою. None, якщо не вдалось."""
    city = _city_slug(target.get("city"), target.get("address"))
    street = _street_tokens(target.get("address"))
    building = _building_slug(target.get("address"))
    if not (city and street and building):
        return None

    # dom.ria використовує зворотний порядок слів вулиці у слагу → пробуємо обидва.
    street_variants = {"-".join(street), "-".join(reversed(street))}
    for street_slug in street_variants:
        url = _DOMRIA_BUILDING_TMPL.format(city=city, street=street_slug, building=building)
        if _verify_catalog(url):
            return url
    return None


def _verify_catalog(url: str) -> bool:
    try:
        response = requests.get(
            url, timeout=_TIMEOUT,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) realtify-autovalue/1.0"},
            allow_redirects=True,
        )
    except Exception:  # noqa: BLE001 — мережевий збій → вважаємо «не знайдено»
        return False
    if response.status_code != 200:
        return False
    return len(set(_LISTING_RE.findall(response.text))) >= _MIN_LISTINGS


def _city_slug(city: Any, address: Any) -> str | None:
    for source in (city, address):
        if not source:
            continue
        text = str(source).lower()
        for name, slug in _CITY_SLUGS.items():
            if name in text:
                return slug
    return None


def _street_tokens(address: Any) -> list[str]:
    """Виділяє слова назви вулиці зі сирого адреса (без міста/будинку/квартири)."""
    if not address:
        return []
    raw = str(address).lower().replace("’", "'")
    raw = re.sub(r"\b(?:кв\.?|квартира|апартаменти?)\s*№?\s*\d+\S*", " ", raw)
    raw = re.sub(r"\bм\.?\s*[а-яіїєґ]+", " ", raw)  # м.Київ
    raw = re.sub(r"\b(?:вул\.?|вулиця|просп\.?|проспект|бульв\.?|бульвар|пров\.?|провулок|"
                 r"набережна|узвіз|площа|шосе|будинок|буд\.?|корпус|корп\.?)\b", " ", raw)
    # Вирізаємо все від першого числа (номер будинку) — далі лише номер/квартира.
    raw = re.split(r"\d", raw, maxsplit=1)[0]
    tokens = [_translit(tok) for tok in re.findall(r"[а-яіїєґ']+", raw) if len(tok) > 1]
    return [t for t in tokens if t]


def _building_slug(address: Any) -> str | None:
    """Номер будинку у форматі dom.ria: «17-К» → «17k», «8» → «8»."""
    if not address:
        return None
    text = str(address).lower()
    # перший «число + опц. літера», що не є номером квартири
    match = re.search(r"(?:будинок|буд\.?|№)?\s*(\d+)\s*[-/]?\s*([а-яіїєґ])?", text)
    if not match:
        return None
    number = match.group(1)
    letter = match.group(2) or ""
    return f"{number}{_translit(letter)}" if letter else number


def _translit(text: str) -> str:
    return "".join(_TRANSLIT.get(ch, ch) for ch in text.lower())
