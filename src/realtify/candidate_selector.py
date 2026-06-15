from __future__ import annotations

import json
import math
import re
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from pydantic import BaseModel, Field

from realtify.models import Comparable


class SelectionError(RuntimeError):
    pass


class CandidateSelectionRecord(BaseModel):
    original_index: int
    selected: bool = False
    selected_rank: int | None = None
    score: float | None = None
    source_url: str
    title: str | None = None
    address: str | None = None
    area_m2: float | None = None
    price_usd: float | None = None
    price_per_m2_usd: float | None = None
    rooms: int | None = None
    source_key: str | None = None
    metrics: dict[str, Any] = Field(default_factory=dict)
    score_reasons: list[str] = Field(default_factory=list)
    rejection_reasons: list[str] = Field(default_factory=list)


@dataclass(frozen=True)
class CandidateSelectionResult:
    required_count: int
    selected_candidates: list[Comparable]
    records: list[CandidateSelectionRecord]
    warnings: list[str]


def select_candidates(
    candidates: list[Comparable],
    *,
    target: dict[str, Any],
    collection_config: dict[str, Any] | None = None,
    required_count: int = 5,
) -> CandidateSelectionResult:
    cfg = _selection_config(collection_config or {}, target)
    if not cfg["enabled"]:
        return _disabled_selection(candidates, required_count)

    records: list[CandidateSelectionRecord] = []
    scored: list[tuple[float, int, Comparable, CandidateSelectionRecord]] = []
    median_price_per_m2 = _median([candidate.price_per_m2_usd for candidate in candidates if candidate.price_per_m2_usd])

    for index, candidate in enumerate(candidates, start=1):
        record = _score_candidate(
            candidate,
            original_index=index,
            target=target,
            cfg=cfg,
            median_price_per_m2=median_price_per_m2,
        )
        records.append(record)
        if record.score is not None and not record.rejection_reasons:
            scored.append((record.score, index, candidate, record))

    scored.sort(key=lambda item: (item[0], item[1]))
    selected_candidates: list[Comparable] = []
    for rank, (_score, _index, candidate, record) in enumerate(scored[:required_count], start=1):
        record.selected = True
        record.selected_rank = rank
        selected_candidates.append(candidate)

    warnings: list[str] = []
    if len(selected_candidates) < required_count:
        warnings.append(f"selected_only_{len(selected_candidates)}_candidate(s)_required_{required_count}")
    if not candidates:
        warnings.append("no_candidates_collected")
    if cfg.get("only_newbuilds") and not cfg.get("require_newbuild_signal"):
        warnings.append("newbuild_filter_assumed_from_discovery_or_source_configuration")

    return CandidateSelectionResult(
        required_count=required_count,
        selected_candidates=selected_candidates,
        records=records,
        warnings=warnings,
    )


def save_selection_result(result: CandidateSelectionResult, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    json_path = output_dir / "candidate_selection.json"
    payload = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "required_count": result.required_count,
        "selected_count": len(result.selected_candidates),
        "warnings": result.warnings,
        "records": [record.model_dump(mode="json") for record in result.records],
        "rejection_summary": dict(_rejection_summary(result.records)),
    }
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    (output_dir / "candidate_selection.md").write_text(_build_selection_markdown(result), encoding="utf-8")
    return json_path


def _score_candidate(
    candidate: Comparable,
    *,
    original_index: int,
    target: dict[str, Any],
    cfg: dict[str, Any],
    median_price_per_m2: float | None,
) -> CandidateSelectionRecord:
    record = CandidateSelectionRecord(
        original_index=original_index,
        source_url=str(candidate.source_url),
        title=candidate.title,
        address=candidate.address,
        area_m2=candidate.area_m2,
        price_usd=candidate.price_usd,
        price_per_m2_usd=candidate.price_per_m2_usd,
        rooms=candidate.rooms,
        source_key=candidate.source_key,
    )
    score = 0.0

    if candidate.transaction_type != str(target.get("transaction_type") or candidate.transaction_type):
        record.rejection_reasons.append("transaction_type_mismatch")
    if candidate.property_type != str(target.get("property_type") or candidate.property_type):
        record.rejection_reasons.append("property_type_mismatch")
    if cfg["require_address"] and not candidate.address:
        record.rejection_reasons.append("address_missing")
    if cfg["require_area"] and candidate.area_m2 is None:
        record.rejection_reasons.append("area_missing")
    if cfg["require_price_usd"] and candidate.price_usd is None:
        record.rejection_reasons.append("price_usd_missing")
    if cfg["require_screenshot"] and not _path_exists(candidate.screenshot_path):
        record.rejection_reasons.append("screenshot_missing")

    city_match = _city_matches(target.get("city"), candidate)
    record.metrics["city_match"] = city_match
    if cfg["require_city_match"] and city_match is False:
        record.rejection_reasons.append("city_mismatch")

    area_delta_pct = _area_delta_pct(target.get("area_m2"), candidate.area_m2)
    record.metrics["area_delta_pct"] = area_delta_pct
    if area_delta_pct is not None:
        if area_delta_pct > cfg["max_area_delta_pct"]:
            record.rejection_reasons.append(f"area_delta_gt_{cfg['max_area_delta_pct']:.0f}_pct")
        score += area_delta_pct * 1.25
        if area_delta_pct > cfg["preferred_area_delta_pct"]:
            score += (area_delta_pct - cfg["preferred_area_delta_pct"]) * 2.0
        record.score_reasons.append(f"area_delta_pct={area_delta_pct:.1f}")

    rooms_delta = _rooms_delta(target.get("rooms"), candidate.rooms)
    record.metrics["rooms_delta"] = rooms_delta
    if rooms_delta is not None:
        if rooms_delta == 0:
            score -= 8.0
            record.score_reasons.append("same_rooms")
        elif cfg["strict_same_rooms"]:
            record.rejection_reasons.append("rooms_mismatch")
        else:
            score += 25.0 + rooms_delta * 12.0
            record.score_reasons.append(f"rooms_delta={rooms_delta}")
    elif _optional_int(target.get("rooms")) is not None:
        score += 15.0
        record.score_reasons.append("candidate_rooms_missing")

    if median_price_per_m2 and candidate.price_per_m2_usd:
        price_delta_pct = abs(candidate.price_per_m2_usd - median_price_per_m2) / median_price_per_m2 * 100
        record.metrics["price_per_m2_delta_from_pool_median_pct"] = round(price_delta_pct, 2)
        score += min(price_delta_pct * 0.35, 35.0)
        record.score_reasons.append(f"price_per_m2_pool_delta_pct={price_delta_pct:.1f}")

    complex_score = _complex_score(target.get("complex_name"), candidate)
    if complex_score:
        score += complex_score
        record.score_reasons.append(f"complex_score={complex_score:.0f}")

    warning_penalty = len(candidate.warnings) * 4.0
    if warning_penalty:
        score += warning_penalty
        record.score_reasons.append(f"candidate_warnings={len(candidate.warnings)}")

    newbuild_signal = _has_newbuild_signal(candidate)
    record.metrics["newbuild_signal"] = newbuild_signal
    if cfg["only_newbuilds"] and cfg["require_newbuild_signal"] and not newbuild_signal:
        record.rejection_reasons.append("newbuild_signal_missing")

    if record.rejection_reasons:
        record.score = None
    else:
        record.score = round(max(score, 0.0), 2)
    return record


def _selection_config(collection_config: dict[str, Any], target: dict[str, Any]) -> dict[str, Any]:
    raw = collection_config.get("selection") or {}
    if not isinstance(raw, dict):
        raise SelectionError("collection.selection must be a YAML object")
    property_type = str(target.get("property_type") or "")
    default_max_area_delta = 75.0 if property_type in {"parking", "commercial", "office", "retail", "warehouse", "house", "land"} else 50.0
    return {
        "enabled": _optional_bool(raw.get("enabled"), default=True),
        "require_address": _optional_bool(raw.get("require_address"), default=True),
        "require_area": _optional_bool(raw.get("require_area"), default=True),
        "require_price_usd": _optional_bool(raw.get("require_price_usd"), default=True),
        "require_screenshot": _optional_bool(raw.get("require_screenshot"), default=True),
        "require_city_match": _optional_bool(raw.get("require_city_match"), default=True),
        "strict_same_rooms": _optional_bool(raw.get("strict_same_rooms"), default=False),
        "only_newbuilds": _optional_bool(collection_config.get("only_newbuilds"), default=False),
        "require_newbuild_signal": _optional_bool(raw.get("require_newbuild_signal"), default=False),
        "preferred_area_delta_pct": _optional_float(raw.get("preferred_area_delta_pct"), default=35.0),
        "max_area_delta_pct": _optional_float(raw.get("max_area_delta_pct"), default=default_max_area_delta),
    }


def _disabled_selection(candidates: list[Comparable], required_count: int) -> CandidateSelectionResult:
    selected = candidates[:required_count]
    records: list[CandidateSelectionRecord] = []
    for index, candidate in enumerate(candidates, start=1):
        selected_rank = index if index <= len(selected) else None
        records.append(
            CandidateSelectionRecord(
                original_index=index,
                selected=selected_rank is not None,
                selected_rank=selected_rank,
                score=0.0 if selected_rank is not None else None,
                source_url=str(candidate.source_url),
                title=candidate.title,
                address=candidate.address,
                area_m2=candidate.area_m2,
                price_usd=candidate.price_usd,
                price_per_m2_usd=candidate.price_per_m2_usd,
                rooms=candidate.rooms,
                source_key=candidate.source_key,
                score_reasons=["selection_disabled"] if selected_rank is not None else [],
            )
        )
    return CandidateSelectionResult(
        required_count=required_count,
        selected_candidates=selected,
        records=records,
        warnings=["candidate_selection_disabled"],
    )


def _city_matches(target_city: Any, candidate: Comparable) -> bool | None:
    aliases = _city_aliases(target_city)
    if not aliases:
        return None
    probe = _normalize_text(" ".join(str(value or "") for value in [candidate.city, candidate.address, candidate.title]))
    if not probe:
        return None
    return any(alias in probe for alias in aliases)


def _city_aliases(value: Any) -> list[str]:
    if value in (None, ""):
        return []
    normalized = _normalize_text(str(value))
    groups = [
        {"київ", "киев", "kyiv", "kiev"},
        {"львів", "львов", "lviv", "lvov"},
        {"одеса", "одесса", "odesa", "odessa"},
        {"дніпро", "днепр", "dnipro", "dnepr"},
        {"харків", "харьков", "kharkiv", "kharkov"},
        {"ужгород", "uzhhorod", "uzhgorod"},
        {"івано франківськ", "ивано франковск", "ivano frankivsk", "ivano frankovsk"},
        {"тернопіль", "тернополь", "ternopil", "ternopol"},
    ]
    for group in groups:
        if normalized in group or any(item in normalized for item in group):
            return sorted(group)
    return [normalized]


def _area_delta_pct(target_area: Any, candidate_area: float | None) -> float | None:
    target = _optional_float(target_area, default=None)
    if target is None or target <= 0 or candidate_area is None:
        return None
    return round(abs(candidate_area - target) / target * 100, 2)


def _rooms_delta(target_rooms: Any, candidate_rooms: int | None) -> int | None:
    target = _optional_int(target_rooms)
    if target is None or candidate_rooms is None:
        return None
    return abs(candidate_rooms - target)


def _complex_score(target_complex: Any, candidate: Comparable) -> float:
    if not target_complex or not candidate.complex_name:
        return 0.0
    target = _normalize_text(str(target_complex))
    candidate_complex = _normalize_text(candidate.complex_name)
    if target and (target in candidate_complex or candidate_complex in target):
        return -6.0
    return 0.0


def _has_newbuild_signal(candidate: Comparable) -> bool:
    probe = _normalize_text(
        " ".join(
            str(value or "")
            for value in [
                candidate.title,
                candidate.address,
                candidate.complex_name,
                candidate.condition,
                candidate.delivery_date,
            ]
        )
    )
    markers = [
        "новобуд",
        "жк ",
        "здача",
        "побудовано",
        "переуступ",
        "забудовник",
        "від забудовника",
    ]
    return any(marker in probe for marker in markers)


def _median(values: list[float | None]) -> float | None:
    cleaned = sorted(float(value) for value in values if value is not None and math.isfinite(float(value)))
    if not cleaned:
        return None
    mid = len(cleaned) // 2
    if len(cleaned) % 2:
        return cleaned[mid]
    return (cleaned[mid - 1] + cleaned[mid]) / 2


def _path_exists(value: Any) -> bool:
    if not value:
        return False
    try:
        return Path(value).exists()
    except (OSError, TypeError, ValueError):
        return False


def _normalize_text(value: str) -> str:
    lowered = value.casefold().replace("ё", "е").replace("’", "'")
    lowered = re.sub(r"[^\w\s'-]+", " ", lowered, flags=re.UNICODE)
    return re.sub(r"\s+", " ", lowered).strip()


def _optional_int(value: Any) -> int | None:
    if value in (None, ""):
        return None
    try:
        return int(float(str(value).replace(",", ".")))
    except (TypeError, ValueError):
        return None


def _optional_float(value: Any, *, default: float | None) -> float | None:
    if value in (None, ""):
        return default
    try:
        return float(str(value).replace(",", "."))
    except (TypeError, ValueError):
        return default


def _optional_bool(value: Any, *, default: bool) -> bool:
    if value in (None, ""):
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().casefold() in {"1", "true", "yes", "y", "так"}


def _rejection_summary(records: list[CandidateSelectionRecord]) -> Counter[str]:
    counter: Counter[str] = Counter()
    for record in records:
        counter.update(record.rejection_reasons)
    return counter


def _build_selection_markdown(result: CandidateSelectionResult) -> str:
    lines = [
        "# Candidate Selection Report",
        "",
        f"Required candidates: {result.required_count}",
        f"Selected candidates: {len(result.selected_candidates)}",
        "",
    ]
    if result.warnings:
        lines.extend(["## Warnings", ""])
        lines.extend(f"- {warning}" for warning in result.warnings)
        lines.append("")

    selected_records = [record for record in result.records if record.selected]
    lines.extend(["## Selected", ""])
    for record in sorted(selected_records, key=lambda item: item.selected_rank or 999):
        lines.extend(_record_lines(record))

    rejected_records = [record for record in result.records if record.rejection_reasons]
    if rejected_records:
        lines.extend(["## Rejected", ""])
        for record in rejected_records:
            lines.extend(_record_lines(record))

    unselected_records = [record for record in result.records if not record.selected and not record.rejection_reasons]
    if unselected_records:
        lines.extend(["## Not Selected", ""])
        for record in sorted(unselected_records, key=lambda item: item.score if item.score is not None else 999999):
            lines.extend(_record_lines(record))
    return "\n".join(lines).strip() + "\n"


def _record_lines(record: CandidateSelectionRecord) -> list[str]:
    title = record.title or record.address or record.source_url
    rank = f" rank={record.selected_rank}" if record.selected_rank else ""
    score = f" score={record.score}" if record.score is not None else ""
    lines = [
        f"### {record.original_index}.{rank}{score} {title}",
        "",
        f"- URL: {record.source_url}",
        f"- Area: {record.area_m2 if record.area_m2 is not None else 'not found'}",
        f"- Price USD: {record.price_usd if record.price_usd is not None else 'not found'}",
        f"- Rooms: {record.rooms if record.rooms is not None else 'not found'}",
    ]
    if record.rejection_reasons:
        lines.append(f"- Rejected: {', '.join(record.rejection_reasons)}")
    if record.score_reasons:
        lines.append(f"- Score reasons: {', '.join(record.score_reasons)}")
    lines.append("")
    return lines
