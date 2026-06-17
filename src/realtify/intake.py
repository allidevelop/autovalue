from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml
from pydantic import BaseModel, Field
from pypdf import PdfReader
from rich.console import Console

from realtify.ocr import OcrTimeoutError, ocr_image
from realtify.paths import PROJECT_ROOT, ensure_output_dir
from realtify.pdf_tools import render_pdf_pages
from realtify.progress import ProgressCallback, emit_progress


class ExtractRecord(BaseModel):
    page: int
    extract_index_number: str | None = None
    extract_formed_at: str | None = None
    registry_object_number: str | None = None
    object_identifier: str | None = None
    object_type: str | None = None
    object_description: str | None = None
    total_area_m2: float | None = None
    living_area_m2: float | None = None
    address_full: str | None = None
    city: str | None = None
    apartment_number: str | None = None
    owners_from_extract: str | None = None
    property_right_type: str | None = None
    warnings: list[str] = Field(default_factory=list)


class TechnicalPassportRecord(BaseModel):
    pages: list[int] = Field(default_factory=list)
    registration_number: str | None = None
    object_identifier: str | None = None
    object_name: str | None = None
    address_full: str | None = None
    city: str | None = None
    apartment_number: str | None = None
    floor_or_level: str | None = None
    total_area_m2: float | None = None
    living_area_m2: float | None = None
    rooms_count: int | None = None
    technical_passport_date: str | None = None
    technical_passport_formed_at: str | None = None
    issuer: str | None = None
    warnings: list[str] = Field(default_factory=list)


class IntakeResult(BaseModel):
    source_pdf: str
    created_at: str
    page_count: int
    pages_text_dir: str
    extracts: list[ExtractRecord]
    technical_passports: list[TechnicalPassportRecord]
    selected_extract: ExtractRecord | None = None
    selected_technical_passport: TechnicalPassportRecord | None = None
    task_yaml_path: str | None = None
    warnings: list[str] = Field(default_factory=list)


class IntakeError(RuntimeError):
    pass


@dataclass(frozen=True)
class IntakeFiles:
    result: IntakeResult
    intake_json: Path
    task_yaml: Path


def write_intake_selection_files(
    *,
    base_result: IntakeResult,
    output_dir: Path,
    selected_extract: ExtractRecord | None,
    selected_technical: TechnicalPassportRecord | None,
    template_path: Path | None = None,
    profile: str = "apartment",
    complex_name: str | None = None,
    links_path: Path | None = None,
) -> IntakeFiles:
    output_dir.mkdir(parents=True, exist_ok=True)
    task_yaml = output_dir / "task.generated.yaml"
    intake_json = output_dir / "intake.json"
    warnings: list[str] = []
    if not base_result.extracts:
        warnings.append("extract_not_found")
    if not base_result.technical_passports:
        warnings.append("technical_passport_not_found")
    if selected_extract is None:
        warnings.append("selected_extract_not_found")
    if selected_technical is None:
        warnings.append("selected_technical_passport_not_found")

    result = base_result.model_copy(
        update={
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "selected_extract": selected_extract,
            "selected_technical_passport": selected_technical,
            "task_yaml_path": str(task_yaml),
            "warnings": warnings,
        }
    )
    intake_json.write_text(
        json.dumps(result.model_dump(mode="json"), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    task_payload = _build_task_yaml(
        pdf_path=Path(base_result.source_pdf),
        selected_extract=selected_extract,
        selected_technical=selected_technical,
        template_path=template_path,
        profile=profile,
        complex_name=complex_name,
        links_path=links_path,
        warnings=warnings,
    )
    task_yaml.write_text(
        yaml.safe_dump(task_payload, allow_unicode=True, sort_keys=False, width=4096),
        encoding="utf-8",
    )
    (output_dir / "intake_summary.md").write_text(_build_intake_summary(result), encoding="utf-8")
    return IntakeFiles(result=result, intake_json=intake_json, task_yaml=task_yaml)


def select_technical_passport_for_extract(
    records: list[TechnicalPassportRecord],
    selected_extract: ExtractRecord | None,
) -> TechnicalPassportRecord | None:
    if selected_extract is None:
        return None
    if selected_extract.apartment_number:
        for record in records:
            if record.apartment_number == selected_extract.apartment_number:
                return record
    if selected_extract.object_identifier:
        for record in records:
            if record.object_identifier == selected_extract.object_identifier:
                return record
    return None


def extract_intake_from_pdf(
    *,
    pdf_path: Path,
    output_dir: Path,
    target_apartment: str | None = None,
    template_path: Path | None = None,
    profile: str = "apartment",
    complex_name: str | None = None,
    links_path: Path | None = None,
    first_page: int | None = None,
    last_page: int | None = None,
    dpi: int = 220,
    force_ocr: bool = False,
    progress: ProgressCallback | None = None,
) -> IntakeFiles:
    output_dir.mkdir(parents=True, exist_ok=True)
    emit_progress(progress, f"PDF intake: reading {pdf_path.name}.")
    page_texts = _extract_or_ocr_pdf(
        pdf_path,
        output_dir=output_dir,
        first_page=first_page,
        last_page=last_page,
        dpi=dpi,
        force_ocr=force_ocr,
        progress=progress,
    )
    emit_progress(progress, "PDF intake: parsing extracts and technical passports.")
    extracts = _backfill_extract_addresses(
        [record for record in (_parse_extract(page, text) for page, text in page_texts.items()) if record]
    )
    technical_passports = _parse_technical_passports(page_texts)
    selected_extract = _select_extract(extracts, target_apartment)
    selected_technical = _select_technical_passport(technical_passports, selected_extract, target_apartment)
    emit_progress(progress, f"PDF intake: found {len(extracts)} extract(s) and {len(technical_passports)} technical passport(s).")

    warnings: list[str] = []
    if not extracts:
        warnings.append("extract_not_found")
    if not technical_passports:
        warnings.append("technical_passport_not_found")
    if selected_extract is None:
        warnings.append("selected_extract_not_found")
    if selected_technical is None:
        warnings.append("selected_technical_passport_not_found")

    task_yaml = output_dir / "task.generated.yaml"
    intake_json = output_dir / "intake.json"

    result = IntakeResult(
        source_pdf=str(pdf_path),
        created_at=datetime.now().isoformat(timespec="seconds"),
        page_count=_pdf_page_count(pdf_path),
        pages_text_dir=str(output_dir / "ocr_text"),
        extracts=extracts,
        technical_passports=technical_passports,
        selected_extract=selected_extract,
        selected_technical_passport=selected_technical,
        task_yaml_path=str(task_yaml),
        warnings=warnings,
    )
    intake_json.write_text(
        json.dumps(result.model_dump(mode="json"), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    task_payload = _build_task_yaml(
        pdf_path=pdf_path,
        selected_extract=selected_extract,
        selected_technical=selected_technical,
        template_path=template_path,
        profile=profile,
        complex_name=complex_name,
        links_path=links_path,
        warnings=warnings,
    )
    task_yaml.write_text(
        yaml.safe_dump(task_payload, allow_unicode=True, sort_keys=False, width=4096),
        encoding="utf-8",
    )
    (output_dir / "intake_summary.md").write_text(_build_intake_summary(result), encoding="utf-8")
    emit_progress(progress, f"PDF intake: wrote intake.json and task.generated.yaml to {output_dir}.")
    return IntakeFiles(result=result, intake_json=intake_json, task_yaml=task_yaml)


def _build_intake_summary(result: IntakeResult) -> str:
    selected_page = result.selected_extract.page if result.selected_extract else None
    selected_apartment = result.selected_extract.apartment_number if result.selected_extract else None
    lines = [
        "# Intake Summary",
        "",
        f"Source PDF: {result.source_pdf}",
        f"Pages: {result.page_count}",
        f"Selected extract page: {selected_page if selected_page is not None else 'not found'}",
        f"Selected apartment/object: {selected_apartment or 'not found'}",
        "",
        "## Extracts Found",
        "",
        "| Page | Apartment | Total area | Living area | Address |",
        "| --- | --- | ---: | ---: | --- |",
    ]
    for record in result.extracts:
        marker = " **SELECTED**" if result.selected_extract and record.page == result.selected_extract.page else ""
        lines.append(
            "| "
            f"{record.page}{marker} | "
            f"{record.apartment_number or ''} | "
            f"{record.total_area_m2 if record.total_area_m2 is not None else ''} | "
            f"{record.living_area_m2 if record.living_area_m2 is not None else ''} | "
            f"{record.address_full or ''} |"
        )

    lines.extend(
        [
            "",
            "## Technical Passports Found",
            "",
            "| Pages | Apartment | Total area | Living area | Rooms |",
            "| --- | --- | ---: | ---: | ---: |",
        ]
    )
    selected_tp_pages = set(result.selected_technical_passport.pages) if result.selected_technical_passport else set()
    for record in result.technical_passports:
        marker = " **SELECTED**" if selected_tp_pages and selected_tp_pages == set(record.pages) else ""
        pages = ", ".join(str(page) for page in record.pages)
        lines.append(
            "| "
            f"{pages}{marker} | "
            f"{record.apartment_number or ''} | "
            f"{record.total_area_m2 if record.total_area_m2 is not None else ''} | "
            f"{record.living_area_m2 if record.living_area_m2 is not None else ''} | "
            f"{record.rooms_count if record.rooms_count is not None else ''} |"
        )
    if result.warnings:
        lines.extend(["", "## Warnings", ""])
        lines.extend(f"- {warning}" for warning in result.warnings)
    return "\n".join(lines).strip() + "\n"


def _extract_or_ocr_pdf(
    pdf_path: Path,
    *,
    output_dir: Path,
    first_page: int | None,
    last_page: int | None,
    dpi: int,
    force_ocr: bool,
    progress: ProgressCallback | None,
) -> dict[int, str]:
    text_dir = output_dir / "ocr_text"
    image_dir = output_dir / "pages"
    text_dir.mkdir(parents=True, exist_ok=True)
    page_count = _pdf_page_count(pdf_path)
    start = first_page or 1
    end = last_page or page_count
    page_texts: dict[int, str] = {}

    text_layer_pages = _extract_text_layer_pages(pdf_path) if not force_ocr else {}
    emit_progress(progress, f"PDF intake: page range {start}-{end}; text layer pages: {len(text_layer_pages)}.")

    for page in range(start, end + 1):
        txt_path = text_dir / f"page_{page:03d}.txt"
        if txt_path.exists() and not force_ocr:
            emit_progress(progress, f"PDF intake: page {page}/{end} using cached text.")
            page_texts[page] = txt_path.read_text(encoding="utf-8")
            continue
        text = text_layer_pages.get(page, "").strip()
        if not text:
            emit_progress(progress, f"PDF intake: page {page}/{end} OCR rendering.")
            images = render_pdf_pages(pdf_path, image_dir, first_page=page, last_page=page, dpi=dpi)
            try:
                text = ocr_image(images[0])
            except OcrTimeoutError as exc:
                emit_progress(progress, f"PDF intake: page {page}/{end} OCR timeout, page skipped: {exc}")
                text = ""
        else:
            emit_progress(progress, f"PDF intake: page {page}/{end} using embedded text.")
        txt_path.write_text(text, encoding="utf-8")
        page_texts[page] = text
    return page_texts


def _extract_text_layer_pages(pdf_path: Path) -> dict[int, str]:
    reader = PdfReader(str(pdf_path))
    pages: dict[int, str] = {}
    for idx, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        if text.strip():
            pages[idx] = text.strip()
    return pages


def _pdf_page_count(pdf_path: Path) -> int:
    return len(PdfReader(str(pdf_path)).pages)


def _parse_extract(page: int, text: str) -> ExtractRecord | None:
    normalized = _normalize_text(text)
    lowered = normalized.casefold()
    if "витяг" not in lowered or "реєстр" not in lowered or "речов" not in lowered:
        return None
    address = _extract_address(normalized)
    address_full = _normalize_address(address)
    object_description = _line_after_label(normalized, "Опис об'єкта") or _line_after_label(normalized, "Опис об’скта")
    total_area, living_area = _parse_areas(normalized)
    owners = _block_after_label(normalized, "Власники", end_markers=["Витяг сформував", "Підпис"])
    record = ExtractRecord(
        page=page,
        extract_index_number=_match_first(
            normalized,
            [
                r"(?:І|I|Ї|П|Т)?ндекс\w*\s+номер\s+витяг\w*\s*:\s*(\d{6,})",
                r"номер\s+витяг\w*\s*:\s*(\d{6,})",
            ],
        ),
        extract_formed_at=_normalize_date_value(
            _match_first(
                normalized,
                [r"Дата[,.]?\s*час\s+формув\w+\s*:\s*([0-9]{2}[.,][0-9]{2}[.,][0-9]{4}\s+[0-9:]{4,8})"],
            )
        ),
        registry_object_number=_match_first(
            normalized,
            [
                r"Реєстраційний\s+номер\s+об['’]єкта\s*\|?\s*([0-9:.\s]{8,})",
                r"номер\s+об['’]єкта\s*\|?\s*([0-9:.\s]{8,})",
            ],
        ),
        object_identifier=_match_first(normalized, [r"(\d{2}\.\d{7}\.\d{7}\.\d{8}\.\d{2}\.\d{4,5}\.\d{2})"]),
        object_type=_line_after_label(normalized, "Тип об'єкта"),
        object_description=object_description,
        total_area_m2=total_area,
        living_area_m2=living_area,
        address_full=address_full,
        city=_parse_city(address_full or normalized),
        apartment_number=_parse_apartment_number(address or normalized),
        owners_from_extract=owners,
        property_right_type=_line_after_label(normalized, "Тип речового права"),
    )
    record.warnings.extend(_missing_warnings(record, ["extract_index_number", "address_full", "total_area_m2"]))
    return record


def _parse_technical_passports(page_texts: dict[int, str]) -> list[TechnicalPassportRecord]:
    records: list[TechnicalPassportRecord] = []
    pages = sorted(page_texts)
    idx = 0
    while idx < len(pages):
        page = pages[idx]
        text = _normalize_text(page_texts[page])
        lowered = text.casefold()
        is_start = (
            ("техн" in lowered and "паспорт" in lowered)
            or "реєстраційний номер у реєстрі будівельної діяльності" in lowered
        )
        if not is_start:
            idx += 1
            continue

        group_pages = [page]
        group_texts = [text]
        j = idx + 1
        while j < len(pages):
            next_page = pages[j]
            next_text = _normalize_text(page_texts[next_page])
            next_lower = next_text.casefold()
            if "витяг" in next_lower and "реєстр" in next_lower and "речов" in next_lower:
                break
            group_pages.append(next_page)
            group_texts.append(next_text)
            j += 1

        records.append(_parse_technical_group(group_pages, "\n".join(group_texts)))
        idx = j
    return records


def _backfill_extract_addresses(records: list[ExtractRecord]) -> list[ExtractRecord]:
    if not records:
        return records
    with_addresses = [record for record in records if record.address_full and record.apartment_number]
    if not with_addresses:
        return records
    patched: list[ExtractRecord] = []
    for record in records:
        if record.address_full or not record.apartment_number:
            patched.append(record)
            continue
        source = min(with_addresses, key=lambda item: abs(item.page - record.page))
        address = _replace_apartment_in_address(source.address_full, record.apartment_number)
        patched.append(
            record.model_copy(
                update={
                    "address_full": address,
                    "city": record.city or source.city or _parse_city(address),
                    "warnings": [*record.warnings, f"address_backfilled_from_page_{source.page}"],
                }
            )
        )
    return patched


def _replace_apartment_in_address(address: str | None, apartment: str) -> str | None:
    if not address:
        return None
    replaced = re.sub(
        r"((?:квартира|квертира|кв\.?/оф\.?)\s*)\d{1,5}",
        rf"\g<1>{apartment}",
        address,
        count=1,
        flags=re.IGNORECASE,
    )
    if replaced == address:
        replaced = f"{address}, квартира {apartment}"
    return _normalize_address(replaced)


def _parse_technical_group(pages: list[int], text: str) -> TechnicalPassportRecord:
    address = _line_after_label(text, "Адреса об'єкта") or _tech_address_from_text(text)
    object_name = _match_first(text, [r"Квартира\s*№\s*(\d+)", r"Квартира\s+N?\s*(\d+)"])
    object_identifier = _match_first(text, [r"(\d{2}\.\d{7}\.\d{7}\.\d{8}\.\d{2}\.\d{4,5}\.\d{2})"])
    total_area, living_area = _parse_areas(text)
    floor = _match_first(text, [r"Поверх\s*:\s*(\d{1,3})"])
    rooms = _rooms_from_explication(text)
    warnings: list[str] = []
    if rooms is None and living_area:
        rooms = 1
        warnings.append("rooms_count_fallback_from_living_area")

    record = TechnicalPassportRecord(
        pages=pages,
        registration_number=_normalize_technical_registration(
            _match_first(
                text,
                [
                    r"Реєстраційний\s+номер\s+(?:у\s+Реєстрі\s+будівельної\s+діяльності|документу)?\s*:?\s*([A-ZА-ЯІЇЄҐ0-9:-]{10,})",
                    r"\b(ТІ01[:0-9-]{10,})",
                ],
            )
        ),
        object_identifier=object_identifier,
        object_name=f"Квартира №{object_name}" if object_name else None,
        address_full=_normalize_address(address),
        city=_parse_city(address),
        apartment_number=object_name or _parse_apartment_number(address or text),
        floor_or_level=floor,
        total_area_m2=total_area,
        living_area_m2=living_area,
        rooms_count=rooms,
        technical_passport_date=_match_first(
            text,
            [r"Дата\s+виготовлення\s+технічного\s+паспорта\s+([^\n\r]+)"],
        ),
        technical_passport_formed_at=_match_first(
            text,
            [r"Дата\s+формування\s+документа\s+([^\n\r]+)", r"Дата\s+створення\s*:\s*([0-9.]+)"],
        ),
        issuer=_match_first(text, [r"ТОВ\s+\"([^\"]*БЮРО[^\"]*)\""]),
        warnings=warnings,
    )
    record.warnings.extend(_missing_warnings(record, ["registration_number", "address_full"]))
    return record


def _build_task_yaml(
    *,
    pdf_path: Path,
    selected_extract: ExtractRecord | None,
    selected_technical: TechnicalPassportRecord | None,
    template_path: Path | None,
    profile: str,
    complex_name: str | None,
    links_path: Path | None,
    warnings: list[str],
) -> dict[str, Any]:
    address = _prefer(selected_extract.address_full if selected_extract else None, selected_technical.address_full if selected_technical else None)
    city = _prefer(selected_extract.city if selected_extract else None, selected_technical.city if selected_technical else None)
    total_area = _prefer(selected_extract.total_area_m2 if selected_extract else None, selected_technical.total_area_m2 if selected_technical else None)
    living_area = _prefer(selected_extract.living_area_m2 if selected_extract else None, selected_technical.living_area_m2 if selected_technical else None)
    rooms = selected_technical.rooms_count if selected_technical else None
    object_type = selected_extract.object_type if selected_extract else None
    apartment_number = _prefer(
        selected_extract.apartment_number if selected_extract else None,
        selected_technical.apartment_number if selected_technical else None,
    )

    return {
        "target": {
            "city": city,
            "address": address,
            "apartment_number": apartment_number,
            "complex_name": complex_name,
            "property_type": _property_type_from_object_type(object_type, profile),
            "transaction_type": "sale",
            "area_m2": total_area,
            "living_area_m2": living_area,
            "rooms": rooms,
            "floor_or_level": selected_technical.floor_or_level if selected_technical else None,
            "location_quality": None,
            "building_class": None,
            "condition": None,
            "delivery_date": "введений" if selected_extract and "закінчений" in (selected_extract.object_description or "").casefold() else None,
        },
        "documents": {
            "mode": "bundle_or_separate",
            "extract_pdf": str(pdf_path),
            "technical_passport_pdf": str(pdf_path),
            "extract_index_number": selected_extract.extract_index_number if selected_extract else None,
            "extract_formed_at": selected_extract.extract_formed_at if selected_extract else None,
            "registry_object_number": selected_extract.registry_object_number if selected_extract else None,
            "owners_from_extract": selected_extract.owners_from_extract if selected_extract else None,
            "technical_passport_registration_number": selected_technical.registration_number if selected_technical else None,
            "technical_passport_date": selected_technical.technical_passport_date if selected_technical else None,
        },
        "template": {
            "profile": profile,
            "path": str(template_path) if template_path else None,
        },
        "collection": {
            "mode": "links" if links_path else "discover",
            "links_path": str(links_path) if links_path else None,
            "sources_config": "config/sources.yaml",
            "required_count": 5,
            "only_newbuilds": True,
            "max_discovered_links": 25,
            "discovery_pages": 1,
            "selection": {
                "enabled": True,
                "require_city_match": True,
                "require_address": True,
                "require_area": True,
                "require_price_usd": True,
                "require_screenshot": True,
                "strict_same_rooms": False,
                "preferred_area_delta_pct": 35,
                "max_area_delta_pct": 50,
            },
            "market_overview_enabled": False,
            "full_page_screenshots": True,
            "report_image_mode": "readable_compressed",
        },
        "adjustments": {
            "bargaining": {
                "enabled": False,
                "note": "Do not overwrite Корегування на торг in MVP.",
            }
        },
        "output": {
            "directory": "outputs",
            "word_format": "docx",
            "monthly_folder": True,
        },
        "intake_warnings": warnings,
    }


def _select_extract(records: list[ExtractRecord], target_apartment: str | None) -> ExtractRecord | None:
    if target_apartment:
        for record in records:
            if record.apartment_number == target_apartment:
                return record
        return None
    return records[0] if records else None


def _select_technical_passport(
    records: list[TechnicalPassportRecord],
    selected_extract: ExtractRecord | None,
    target_apartment: str | None,
) -> TechnicalPassportRecord | None:
    apartment = target_apartment or (selected_extract.apartment_number if selected_extract else None)
    if apartment:
        for record in records:
            if record.apartment_number == apartment:
                return record
    if selected_extract and selected_extract.object_identifier:
        for record in records:
            if record.object_identifier == selected_extract.object_identifier:
                return record
    return records[0] if records else None


def _rooms_from_explication(text: str) -> int | None:
    # Works when OCR keeps rows with room names and living area values. Scanned tables often need manual review.
    room_lines = []
    for line in text.splitlines():
        lowered = line.casefold()
        if any(word in lowered for word in ["житлова", "кімната", "спальня"]):
            if re.search(r"\d+[.,]\d+", line):
                room_lines.append(line)
    return len(room_lines) or None


def _parse_areas(text: str) -> tuple[float | None, float | None]:
    total = _match_decimal(
        text,
        [
            r"Загальна\s+\S*п?лоща\s*\(кв\.?м\)\s*[:;]\s*(\d+[.,]\d+)",
            r"Загальна\s+[^\n\r]{0,40}?\(кв\.?м\)\s*[:;]\s*(\d+[.,]\d+)",
            r"Загальна\s+площа\s+приміщень[^\d]*(\d+[.,]\d+)",
        ],
    )
    living = _match_decimal(
        text,
        [
            r"житлова\s+\S*п?лоща\s*\(кв\.?м\)\s*[:;]\s*(\d+[.,]\d+)",
            r"Житлова\s+площа\s+приміщень[^\d]*(\d+[.,]\d+)",
        ],
    )
    if total is None:
        total = _match_decimal(
            text,
            [
                r"Загальна\s+площа\s*\([^)]{0,12}\)\s*[:;]\s*(\d+[.,]\d+)",
                r"Загальна\s+площа[^\n\r:]{0,24}\s*[:;]\s*(\d+[.,]\d+)",
                r"Загальна\s+\S{3,16}\s*\([^)]{0,12}\)\s*[:;]\s*(\d+(?:[.,]\d+)?)",
            ],
        )
    if living is None:
        living = _match_decimal(
            text,
            [
                r"житлова\s+площа\s*\([^)]{0,12}\)\s*[:;]\s*(\d+[.,]\d+)",
                r"житлова\s+площа[^\n\r:]{0,24}\s*[:;]\s*(\d+[.,]\d+)",
                r"житл\w+\s+площа\s*\([^)]{0,12}\)\s*[:;]\s*(\d+(?:[.,]\d+)?)",
            ],
        )
    if total is not None and living is not None and living > total:
        scaled_living = living / 10
        living = scaled_living if scaled_living <= total else None
    return total, living


def _line_after_label(text: str, label: str) -> str | None:
    escaped = re.escape(label).replace("\\'", "['’`]")
    match = re.search(rf"{escaped}\s*[:|]?\s*(.+)", text, re.IGNORECASE)
    if not match:
        return None
    value = _clean_line(match.group(1))
    if not value and match.end() < len(text):
        tail = text[match.end() :].splitlines()
        value = _clean_line(tail[0]) if tail else ""
    return value or None


def _block_after_label(text: str, label: str, *, end_markers: list[str]) -> str | None:
    match = re.search(rf"{re.escape(label)}\s*[:|]?\s*(.+)", text, re.IGNORECASE | re.DOTALL)
    if not match:
        return None
    block = match.group(1)
    for marker in end_markers:
        idx = block.casefold().find(marker.casefold())
        if idx >= 0:
            block = block[:idx]
    lines = [_clean_line(line) for line in block.splitlines()]
    lines = [line for line in lines if line]
    return " ".join(lines[:8]) or None


def _tech_address_from_text(text: str) -> str | None:
    match = re.search(r"(м\.\s*Київ[^\n]{0,120}(?:кв\.?/оф\.?\s*\d+|квартира\s*\d+))", text, re.IGNORECASE)
    return _clean_line(match.group(1)) if match else None


def _extract_address(text: str) -> str | None:
    for label in ["Адреса", "Аярес", "Адрсс", "Арес"]:
        labeled = _line_after_label(text, label)
        if _looks_like_full_address(labeled):
            return labeled
    fuzzy_labeled = _address_after_fuzzy_label(text)
    if fuzzy_labeled:
        return fuzzy_labeled
    return _extract_address_from_text(text)


def _address_after_fuzzy_label(text: str) -> str | None:
    lines = text.splitlines()
    for index, line in enumerate(lines):
        label = _clean_line(line).casefold()
        if not re.fullmatch(r"(?:адреса?|аярес|а.?рес)", label):
            continue
        parts: list[str] = []
        for next_line in lines[index + 1 : index + 12]:
            cleaned = _clean_line(next_line)
            if not cleaned:
                continue
            lowered = cleaned.casefold()
            if any(
                marker in lowered
                for marker in [
                    "актуальна інформація",
                    "кауальна інформація",
                    "номер відомостей",
                    "речове право",
                    "власники",
                ]
            ):
                break
            parts.append(cleaned)
            candidate = _clean_line(" ".join(parts))
            if _looks_like_full_address(candidate):
                return candidate
    return None


def _extract_address_from_text(text: str) -> str | None:
    patterns = [
        r"(м\.\s*Київ[^\n\r]{0,180}(?:квартира|квертира|кнартира|квазтира|кв\.?/оф\.?)\s*\d{1,5})",
        r"(Київ[^\n\r]{0,180}(?:квартира|квертира|кнартира|квазтира|кв\.?/оф\.?)\s*\d{1,5})",
        r"((?:лов|м\.)?\s*,?\s*Д[^\n\r]{0,45}набережна[^\n\r]{0,180}(?:квартира|квертира|кнартира|квазтира|кв\.?/оф\.?)\s*\d{1,5})",
        r"((?:м\.?\s*Київ|Київ|хо\s*він|лов)?\s*,?\s*Д[^\n\r]{0,80}набережна[^\n\r]{0,180}(?:квартира|квертира|кнартира|квазтира|кв\.?/оф\.?)\s*\d{1,5})",
    ]
    for search_text in [text, re.sub(r"\s+", " ", text)]:
        for pattern in patterns:
            match = re.search(pattern, search_text, re.IGNORECASE)
            if match:
                return _clean_line(match.group(1))
    return None


def _looks_like_address(value: str | None) -> bool:
    if not value:
        return False
    lowered = value.casefold()
    return (
        _parse_city(value) is not None
        or _parse_apartment_number(value) is not None
        or "набереж" in lowered
        or "будин" in lowered
    )


def _looks_like_full_address(value: str | None) -> bool:
    normalized = _normalize_address(value) or value
    if not normalized or not _looks_like_address(normalized):
        return False
    lowered = normalized.casefold()
    return _parse_apartment_number(normalized) is not None or "будин" in lowered


def _normalize_text(text: str) -> str:
    replacements = {
        "\xa0": " ",
        "  ": " ",
        "будинох": "будинок",
        "будилох": "будинок",
        "будилок": "будинок",
        "будипск": "будинок",
        "будинск": "будинок",
        "будииок": "будинок",
        "Дпіпровська": "Дніпровська",
        "Диїпровська": "Дніпровська",
        "Дипровська": "Дніпровська",
        "Диіпровем": "Дніпровська",
        "Дніпровем": "Дніпровська",
        "Дийпровська": "Дніпровська",
        "Диіпрозська": "Дніпровська",
        "Днілровська": "Дніпровська",
        "аборожно": "набережна",
        "пабережна": "набережна",
        "пабережица": "набережна",
        "набережица": "набережна",
        "квертира": "квартира",
        "хо він": "м.Київ",
        "заКиїв": "м.Київ",
        "Хоїв": "Київ",
        "Каїв": "Київ",
    }
    normalized = text
    for old, new in replacements.items():
        normalized = normalized.replace(old, new)
    normalized = re.sub(r"[ \t]+", " ", normalized)
    return normalized.strip()


def _normalize_address(address: str | None) -> str | None:
    if not address:
        return None
    value = _clean_line(address)
    value = value.replace("Дийпровська", "Дніпровська")
    value = value.replace("Диіпровем", "Дніпровська").replace("Дніпровем", "Дніпровська")
    value = value.replace("аборожно", "набережна")
    value = value.replace("пабережна", "набережна").replace("пабережица", "набережна").replace("набережица", "набережна")
    value = value.replace("будилох", "будинок").replace("будилок", "будинок").replace("будииок", "будинок").replace("будипск", "будинок").replace("будинск", "будинок")
    value = value.replace("хо він", "м.Київ")
    value = value.replace("заКиїв", "м.Київ")
    value = value.replace("17- | К", "17-К").replace("17- К", "17-К").replace("17-K", "17-К")
    value = value.replace("І7-К", "17-К").replace("I7-К", "17-К").replace("l7-К", "17-К")
    value = value.replace("кв./оф.", "квартира")
    value = re.sub(r"\b(?:квертира|кнартира|квазтира)\b", "квартира", value, flags=re.IGNORECASE)
    if re.match(r"(?i)^\s*лов\s*,\s*Дніпровська", value):
        value = re.sub(r"(?i)^\s*лов\s*,\s*", "м.Київ, ", value, count=1)
    if "Дніпровська набережна" in value and not _parse_city(value):
        value = f"м.Київ, {value.lstrip(' ,')}"
    value = re.sub(r"(набережна)\s+(будинок)", r"\1, \2", value, flags=re.IGNORECASE)
    value = value.replace("інший ", "")
    value = value.replace("(місце ", "")
    value = re.sub(r"\s*,\s*", ", ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip(" ,")


def _normalize_date_value(value: str | None) -> str | None:
    if not value:
        return None
    return re.sub(r"(?<=\d{2}\.\d{2}),(\d{4})", r".\1", value)


def _normalize_technical_registration(value: str | None) -> str | None:
    if not value:
        return None
    return value.replace("T101", "ТІ01").replace("Т101", "ТІ01")


def _parse_city(address: str | None) -> str | None:
    if not address:
        return None
    if "Київ" in address:
        return "Київ"
    if "Дніпровська набережна" in address:
        return "Київ"
    if "Львів" in address:
        return "Львів"
    return None


def _parse_apartment_number(text: str) -> str | None:
    match = re.search(r"(?:квартира|квертира|кнартира|квазтира|кв\.?/оф\.?|Квартира\s*№)\s*([0-9]{1,5})", text, re.IGNORECASE)
    return match.group(1) if match else None


def _property_type_from_object_type(object_type: str | None, fallback: str) -> str:
    if object_type and "квартира" in object_type.casefold():
        return "apartment"
    return fallback


def _match_first(text: str, patterns: list[str]) -> str | None:
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if match:
            return _clean_line(match.group(1))
    return None


def _match_decimal(text: str, patterns: list[str]) -> float | None:
    value = _match_first(text, patterns)
    if not value:
        return None
    try:
        return float(value.replace(",", "."))
    except ValueError:
        return None


def _prefer(*values: Any) -> Any:
    for value in values:
        if value not in (None, ""):
            return value
    return None


def _missing_warnings(record: BaseModel, fields: list[str]) -> list[str]:
    return [f"{field}_not_found" for field in fields if getattr(record, field) in (None, "")]


def _clean_line(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("|", " ")).strip(" :;,.")


def _resolve_path(path: Path) -> Path:
    return path if path.is_absolute() else PROJECT_ROOT / path


def _default_output_dir() -> Path:
    return ensure_output_dir(datetime.now().strftime("%Y%m%d_%H%M%S_intake"))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Extract valuation intake data from scanned PDF documents.")
    parser.add_argument("--pdf", type=Path, required=True, help="Input extract/technical-passport PDF.")
    parser.add_argument("--out", type=Path, default=None, help="Output directory.")
    parser.add_argument("--apartment", default=None, help="Target apartment number in a bundled PDF.")
    parser.add_argument("--template", type=Path, default=None, help="Excel template path to write into generated task YAML.")
    parser.add_argument("--profile", default="apartment", help="Template profile name.")
    parser.add_argument("--complex-name", default=None)
    parser.add_argument("--links", type=Path, default=None, help="Optional links file path for generated task YAML.")
    parser.add_argument("--first-page", type=int, default=None)
    parser.add_argument("--last-page", type=int, default=None)
    parser.add_argument("--dpi", type=int, default=220)
    parser.add_argument("--force-ocr", action="store_true")
    args = parser.parse_args(argv)

    console = Console()
    try:
        files = extract_intake_from_pdf(
            pdf_path=_resolve_path(args.pdf),
            output_dir=_resolve_path(args.out) if args.out else _default_output_dir(),
            target_apartment=args.apartment,
            template_path=_resolve_path(args.template) if args.template else None,
            profile=args.profile,
            complex_name=args.complex_name,
            links_path=_resolve_path(args.links) if args.links else None,
            first_page=args.first_page,
            last_page=args.last_page,
            dpi=args.dpi,
            force_ocr=args.force_ocr,
        )
    except Exception as exc:
        console.print(f"[red]Intake extraction failed:[/red] {exc}")
        return 1

    selected = files.result.selected_extract
    selected_tp = files.result.selected_technical_passport
    console.print("[green]Intake extraction complete[/green]")
    console.print(f"Extracts found: {len(files.result.extracts)}")
    console.print(f"Technical passports found: {len(files.result.technical_passports)}")
    console.print(f"Selected apartment: {selected.apartment_number if selected else 'not found'}")
    console.print(f"Selected tech passport: {selected_tp.apartment_number if selected_tp else 'not found'}")
    if files.result.warnings:
        for warning in files.result.warnings:
            console.print(f"[yellow]Warning:[/yellow] {warning}")
    console.print(f"Intake JSON: {files.intake_json}")
    console.print(f"Task YAML: {files.task_yaml}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
