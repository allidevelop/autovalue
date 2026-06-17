from __future__ import annotations

import argparse
import json
import os
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml
from rich.console import Console

from realtify import analog_cache
from realtify.candidate_selector import CandidateSelectionResult, save_selection_result, select_candidates
from realtify.collect_from_links import CollectionResult, collect_from_links, read_links, save_collection_result
from realtify.discover_links import DiscoveryResult, discover_links_for_task, save_discovery_result
from realtify.fill_template import FillResult, fill_excel_template, load_template_profile
from realtify.models import PropertyType, TransactionType
from realtify.paths import PROJECT_ROOT, RESOURCE_ROOT, ensure_output_dir
from realtify.progress import ProgressCallback, emit_progress
from realtify.source_config import load_sources_config
from realtify import valuation_register
from realtify.valuation_date import ValuationContext, resolve_valuation_context


class WorkflowError(RuntimeError):
    pass


@dataclass(frozen=True)
class WorkflowResult:
    output_dir: Path
    collection: CollectionResult
    excel: FillResult | None
    report_path: Path
    discovery: DiscoveryResult | None = None
    raw_collection: CollectionResult | None = None
    selection: CandidateSelectionResult | None = None


def run_excel_workflow(
    *,
    task_path: Path,
    links_path: Path | None = None,
    output_dir: Path | None = None,
    required_count: int | None = None,
    allow_less: bool = False,
    allow_incomplete: bool = False,
    visible: bool = False,
    progress: ProgressCallback | None = None,
) -> WorkflowResult:
    emit_progress(progress, "Excel workflow: читаю task.generated.yaml та профіль шаблону.")
    task_file = _resolve_path(task_path)
    task = _load_yaml(task_file)
    target = _section(task, "target")
    template = _section(task, "template")
    collection_cfg = _section(task, "collection")

    profile_name = template.get("profile")
    if not profile_name:
        raise WorkflowError("task.template.profile is required")
    profile_path = RESOURCE_ROOT / "config" / "template_profiles" / f"{profile_name}.yaml"
    profile = load_template_profile(profile_path)

    template_path = _optional_path(template.get("path"))
    if not template_path:
        raise WorkflowError("task.template.path is required")
    template_path = _resolve_path(template_path)

    sources_path = _optional_path(collection_cfg.get("sources_config"))
    sources_config = load_sources_config(_resolve_resource_path(sources_path) if sources_path else None)

    required = (
        required_count
        or _optional_int(collection_cfg.get("required_count"))
        or sources_config.defaults.required_count
    )
    property_type = _typed_property_type(target.get("property_type") or profile.profile)
    transaction_type = _typed_transaction_type(target.get("transaction_type") or "sale")
    out_dir = _resolve_path(output_dir) if output_dir else _default_workflow_output_dir(target)
    out_dir.mkdir(parents=True, exist_ok=True)

    # ── Дата оцінки (реєстр клієнта) + курс НБУ на цю дату ──
    valuation = resolve_valuation_context(task)
    for note in valuation.notes:
        emit_progress(progress, note)

    # ── Кеш аналогів за адресою: для повторюваних адрес пропускаємо пошук ──
    cache_key = analog_cache.address_key(
        city=target.get("city"),
        address=target.get("address"),
        property_type=str(property_type),
        complex_name=target.get("complex_name"),
    )
    manual_links = links_path or _optional_path(collection_cfg.get("links_path"))
    cached = analog_cache.lookup(cache_key, out_dir) if (cache_key and not manual_links) else None

    links_file: Path | None = manual_links
    discovery: DiscoveryResult | None = None
    collection: CollectionResult | None = None
    selection: CandidateSelectionResult | None = None

    if cached and len(cached) >= required:
        emit_progress(
            progress,
            "Кеш аналогів: адресу знайдено в базі — пропускаю пошук, "
            f"беру {len(cached)} збережених аналогів.",
        )
        selected_collection = CollectionResult(
            output_dir=out_dir,
            candidates=cached[:required] if required else cached,
            errors=[],
        )
        links_file = out_dir / "candidates.json"
        save_collection_result(
            selected_collection,
            links=[str(c.source_url) for c in selected_collection.candidates],
            candidates_filename="candidates.json",
            report_filename="selected_candidates_report.md",
        )
    else:
        if links_file:
            links_file = _resolve_path(links_file)
            links = read_links(links_file)
            emit_progress(progress, f"Використовую ручний список посилань: {links_file} ({len(links)} URL).")
        else:
            emit_progress(progress, "Автоматичний пошук аналогів: старт.")
            discovery = discover_links_for_task(
                task=task,
                output_dir=out_dir,
                sources_config=sources_config,
                required_count=required,
                max_links=_optional_int(collection_cfg.get("max_discovered_links")),
                pages_per_source=_optional_int(collection_cfg.get("discovery_pages")),
                progress=progress,
            )
            links_file = save_discovery_result(discovery)
            links = [str(link.url) for link in discovery.links]
            emit_progress(progress, f"Автоматичний пошук аналогів: збережено {len(links)} URL у {links_file.name}.")

        emit_progress(progress, "Збір даних та full-page screenshots по знайдених оголошеннях: старт.")
        collection = collect_from_links(
            links,
            output_dir=out_dir,
            sources_config=sources_config,
            property_type=property_type,
            transaction_type=transaction_type,
            progress=progress,
        )
        save_collection_result(
            collection,
            links=links,
            candidates_filename="collected_candidates.json",
            report_filename="collection_report.md",
        )
        emit_progress(progress, f"Сирий пул аналогів збережено: {len(collection.candidates)} кандидатів, {len(collection.errors)} помилок.")
        emit_progress(progress, "Відбір 5 найкращих аналогів: старт.")
        selection = select_candidates(
            collection.candidates,
            target=target,
            collection_config=collection_cfg,
            required_count=required,
        )
        save_selection_result(selection, out_dir)
        emit_progress(progress, f"Відбір завершено: обрано {len(selection.selected_candidates)} з {len(collection.candidates)} кандидатів.")
        if any(w.startswith("analogs_not_same_complex") for w in selection.warnings):
            emit_progress(
                progress,
                "УВАГА: жоден аналог не з того ж ЖК/будинку — оцінка орієнтовна. "
                "Вкажіть «Назва ЖК» або додайте ручні посилання на аналоги цього будинку.",
            )
        selected_collection = CollectionResult(
            output_dir=out_dir,
            candidates=selection.selected_candidates,
            errors=collection.errors,
        )
        save_collection_result(
            selected_collection,
            links=[str(candidate.source_url) for candidate in selection.selected_candidates],
            candidates_filename="candidates.json",
            report_filename="selected_candidates_report.md",
        )
        # Кешуємо підібрані аналоги ТІЛЬКИ якщо вони достовірні: того ж ЖК/будинку
        # або задані вручну. Інакше не розмножуємо city-wide «сміття» по кешу.
        selected_same = (
            sum(
                1
                for rec in selection.records
                if rec.selected and rec.metrics.get("same_complex_or_address") is True
            )
            if selection
            else 0
        )
        cache_trustworthy = bool(manual_links) or selected_same >= 1
        if cache_key and selected_collection.candidates and cache_trustworthy:
            try:
                analog_cache.save(
                    cache_key,
                    city=target.get("city"),
                    address=target.get("address"),
                    property_type=str(property_type),
                    complex_name=target.get("complex_name"),
                    candidates=selected_collection.candidates,
                )
                emit_progress(
                    progress,
                    f"Кеш аналогів: збережено {len(selected_collection.candidates)} аналогів для адреси.",
                )
            except Exception as exc:  # noqa: BLE001 — кеш не повинен валити основний потік
                emit_progress(progress, f"Кеш аналогів: не вдалося зберегти ({exc}).")
        elif cache_key and selected_collection.candidates:
            emit_progress(
                progress,
                "Кеш аналогів: НЕ зберігаю — аналоги не з того ж ЖК/будинку "
                "(уникаю кешування неточних).",
            )

    excel_result: FillResult | None = None
    fill_error: str | None = None
    try:
        emit_progress(progress, "Заповнення Excel-шаблону: старт.")
        excel_result = fill_excel_template(
            template_path=template_path,
            profile=profile,
            candidates=selected_collection.candidates,
            output_path=out_dir / f"{profile.profile}_filled.xls",
            target=target,
            required_count=required,
            allow_less=allow_less,
            allow_incomplete=allow_incomplete,
            visible=visible,
            nbu_rate=valuation.nbu_rate,
        )
        emit_progress(progress, f"Excel-шаблон заповнено: {excel_result.output_path}")
        _writeback_estimate_to_register(task, valuation, excel_result, progress=progress)
    except Exception as exc:
        fill_error = str(exc)
        if not allow_less and len(selected_collection.candidates) < required:
            _write_workflow_report(
                out_dir,
                links_file=links_file,
                required_count=required,
                collection=selected_collection,
                raw_collection=collection,
                excel_result=None,
                fill_error=fill_error,
                discovery=discovery,
                selection=selection,
            )
            raise WorkflowError(fill_error) from exc
        if not allow_incomplete:
            _write_workflow_report(
                out_dir,
                links_file=links_file,
                required_count=required,
                collection=selected_collection,
                raw_collection=collection,
                excel_result=None,
                fill_error=fill_error,
                discovery=discovery,
                selection=selection,
            )
            raise WorkflowError(fill_error) from exc
        raise

    report_path = _write_workflow_report(
        out_dir,
        links_file=links_file,
        required_count=required,
        collection=selected_collection,
        raw_collection=collection,
        excel_result=excel_result,
        fill_error=fill_error,
        discovery=discovery,
        selection=selection,
    )
    emit_progress(progress, f"Excel workflow завершено. Технічний звіт: {report_path}")
    return WorkflowResult(
        output_dir=out_dir,
        collection=selected_collection,
        excel=excel_result,
        report_path=report_path,
        discovery=discovery,
        raw_collection=collection,
        selection=selection,
    )


def _write_workflow_report(
    output_dir: Path,
    *,
    links_file: Path,
    required_count: int,
    collection: CollectionResult,
    raw_collection: CollectionResult | None,
    excel_result: FillResult | None,
    fill_error: str | None,
    discovery: DiscoveryResult | None = None,
    selection: CandidateSelectionResult | None = None,
) -> Path:
    report_path = output_dir / "report.md"
    candidates_json = output_dir / "candidates.json"
    raw_count = len(raw_collection.candidates) if raw_collection else len(collection.candidates)
    lines = [
        "# Excel Workflow Report",
        "",
        f"Created: {datetime.now().isoformat(timespec='seconds')}",
        f"Links file: {links_file}",
        f"Required candidates: {required_count}",
        f"Candidates collected: {raw_count}",
        f"Candidates selected: {len(collection.candidates)}",
        f"Errors: {len(collection.errors)}",
        f"Candidates JSON: {candidates_json}",
        f"Collected candidates JSON: {output_dir / 'collected_candidates.json'}",
        f"Candidate selection JSON: {output_dir / 'candidate_selection.json'}",
        f"Excel output: {excel_result.output_path if excel_result else 'not created'}",
    ]
    if discovery:
        lines.extend(
            [
                f"Discovery JSON: {output_dir / 'discovery.json'}",
                f"Discovered links: {len(discovery.links)}",
                f"Discovery source pages: {len(discovery.source_pages)}",
            ]
        )
        if discovery.warnings:
            lines.extend(["", "## Discovery Warnings", ""])
            lines.extend(f"- {warning}" for warning in discovery.warnings)
    if excel_result and excel_result.warnings:
        lines.extend(["", "## Excel Warnings", ""])
        lines.extend(f"- {warning}" for warning in excel_result.warnings)
    if selection and selection.warnings:
        lines.extend(["", "## Selection Warnings", ""])
        lines.extend(f"- {warning}" for warning in selection.warnings)
    if fill_error:
        lines.extend(["", "## Excel Error", "", fill_error])

    lines.extend(["", "## Selected Candidates", ""])
    for idx, candidate in enumerate(collection.candidates, start=1):
        lines.extend(
            [
                f"### {idx}. {candidate.title or candidate.source_name or 'Listing'}",
                "",
                f"- Source: {candidate.source_name or 'unknown'}",
                f"- URL: {candidate.source_url}",
                f"- Address: {candidate.address or 'not found'}",
                f"- Area: {candidate.area_m2 if candidate.area_m2 is not None else 'not found'}",
                f"- Price USD: {candidate.price_usd if candidate.price_usd is not None else 'not found'}",
                f"- Screenshot: {candidate.screenshot_path or 'not saved'}",
                f"- Report image: {candidate.report_image_path or 'not saved'}",
                f"- Warnings: {', '.join(candidate.warnings) if candidate.warnings else 'none'}",
                "",
            ]
        )

    if selection:
        rejected = [record for record in selection.records if record.rejection_reasons]
        if rejected:
            lines.extend(["## Rejected By Selection", ""])
            for record in rejected[:30]:
                lines.append(
                    f"- {record.source_url}: {', '.join(record.rejection_reasons)}"
                )

    if collection.errors:
        lines.extend(["## Rejected Links", ""])
        for error in collection.errors:
            extra = f" screenshot={error['screenshot_path']}" if error.get("screenshot_path") else ""
            lines.append(f"- `{error['source']}` {error['url']}: {error['error']}{extra}")

    report_path.write_text("\n".join(lines).strip() + "\n", encoding="utf-8")
    return report_path


def _load_yaml(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    if not isinstance(data, dict):
        raise WorkflowError(f"{path} must contain a YAML object")
    return data


def _writeback_estimate_to_register(
    task: dict[str, Any],
    valuation: ValuationContext,
    excel_result: FillResult,
    *,
    progress: ProgressCallback | None,
) -> None:
    """Записує посчитану оцінку в окремі колонки реєстру (рядок зматченого об'єкта)."""
    entry = valuation.matched_entry
    if entry is None or not _writeback_enabled(task):
        return
    if not valuation.apartment_verified:
        # № квартири не підтверджено точно — не пишемо, щоб не потрапити в чужий рядок.
        return
    register_path = valuation_register.register_path_from_task(task)
    if not register_path:
        return
    estimate = _read_market_value(excel_result)
    if estimate is None:
        emit_progress(progress, "Реєстр: оцінку не записано — не вдалося прочитати market_value з розрахунку.")
        return
    try:
        written = valuation_register.write_estimate(
            register_path,
            entry,
            estimate_uah=estimate,
            nbu_rate=valuation.nbu_rate,
            calc_date=datetime.now().date(),
        )
    except Exception as exc:  # noqa: BLE001 — запис у реєстр не повинен валити оцінку
        emit_progress(progress, f"Реєстр: не вдалося записати оцінку ({exc}).")
        return
    if written:
        amount = f"{estimate:,.0f}".replace(",", " ")
        emit_progress(
            progress,
            f"Реєстр: записано оцінку {amount} грн у рядок кв. {entry.apartment or '?'} "
            f"(аркуш «{entry.sheet}», р.{entry.row}); клієнтську «Ціна продажу» не змінено.",
        )


def _writeback_enabled(task: dict[str, Any]) -> bool:
    if os.environ.get("REALTIFY_REGISTER_WRITEBACK", "1").strip().lower() in ("0", "false", "no", "off"):
        return False
    valuation_cfg = task.get("valuation") if isinstance(task.get("valuation"), dict) else {}
    flag = valuation_cfg.get("writeback")
    return True if flag is None else bool(flag)


def _read_market_value(excel_result: FillResult) -> float | None:
    meta = excel_result.metadata_path
    if not meta:
        return None
    try:
        with open(meta, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
        summary = payload.get("summary") or {}
        value = summary.get("market_value_uah_rounded")
        if value is None:
            value = summary.get("market_value_uah")
        return float(value) if value is not None else None
    except Exception:  # noqa: BLE001 — відсутній/битий сайдкар → нічого писати
        return None


def _section(task: dict[str, Any], name: str) -> dict[str, Any]:
    value = task.get(name) or {}
    if not isinstance(value, dict):
        raise WorkflowError(f"task.{name} must be a YAML object")
    return value


def _resolve_path(path: Path) -> Path:
    return path if path.is_absolute() else PROJECT_ROOT / path


def _resolve_resource_path(path: Path) -> Path:
    if path.is_absolute():
        return path
    project_candidate = PROJECT_ROOT / path
    if project_candidate.exists():
        return project_candidate
    return RESOURCE_ROOT / path


def _optional_path(value: Any) -> Path | None:
    if value in (None, ""):
        return None
    return Path(str(value))


def _optional_int(value: Any) -> int | None:
    if value in (None, ""):
        return None
    return int(value)


def _typed_property_type(value: Any) -> PropertyType:
    if value not in PropertyType.__args__:
        raise WorkflowError(f"Unsupported property_type: {value}")
    return value


def _typed_transaction_type(value: Any) -> TransactionType:
    if value not in TransactionType.__args__:
        raise WorkflowError(f"Unsupported transaction_type: {value}")
    return value


def _default_workflow_output_dir(target: dict[str, Any]) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    address = str(target.get("address") or target.get("complex_name") or "workflow")
    slug = re.sub(r"[^A-Za-z0-9_-]+", "-", address).strip("-").lower()
    if not slug:
        slug = "workflow"
    return ensure_output_dir(f"{timestamp}_{slug[:40]}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run discovery/links -> screenshots -> candidates -> filled Excel workflow.")
    parser.add_argument("--task", type=Path, required=True, help="Task YAML file.")
    parser.add_argument("--links", type=Path, default=None, help="UTF-8 text file with one URL per line.")
    parser.add_argument("--out", type=Path, default=None, help="Output directory.")
    parser.add_argument("--required-count", type=int, default=None)
    parser.add_argument("--allow-less", action="store_true", help="Allow fewer candidates than required.")
    parser.add_argument("--allow-incomplete", action="store_true", help="Allow missing critical candidate fields.")
    parser.add_argument("--visible", action="store_true", help="Show Excel while filling.")
    args = parser.parse_args(argv)

    console = Console()
    try:
        result = run_excel_workflow(
            task_path=args.task,
            links_path=args.links,
            output_dir=args.out,
            required_count=args.required_count,
            allow_less=args.allow_less,
            allow_incomplete=args.allow_incomplete,
            visible=args.visible,
        )
    except Exception as exc:
        console.print(f"[red]Workflow failed:[/red] {exc}")
        return 1

    console.print(f"[green]Workflow complete[/green]")
    raw_count = len(result.raw_collection.candidates) if result.raw_collection else len(result.collection.candidates)
    console.print(f"Candidates collected: {raw_count}")
    console.print(f"Candidates selected: {len(result.collection.candidates)}")
    console.print(f"Errors: {len(result.collection.errors)}")
    console.print(f"Excel: {result.excel.output_path if result.excel else 'not created'}")
    console.print(f"Report: {result.report_path}")
    console.print(f"Output: {result.output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
