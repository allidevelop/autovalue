from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from rich.console import Console

from realtify.full_workflow import FullWorkflowError
from realtify.intake import (
    ExtractRecord,
    IntakeFiles,
    extract_intake_from_pdf,
    select_technical_passport_for_extract,
    write_intake_selection_files,
)
from realtify.paths import PROJECT_ROOT, ensure_output_dir
from realtify.progress import ProgressCallback, emit_progress
from realtify.report_generator import generate_word_report
from realtify.report_validator import ReportValidationResult, validate_report
from realtify.workflow import WorkflowResult, run_excel_workflow


@dataclass(frozen=True)
class BatchObjectResult:
    output_dir: Path
    extract_page: int | None
    apartment: str | None
    intake: IntakeFiles | None
    excel_workflow: WorkflowResult | None
    word_report_path: Path | None
    validation: ReportValidationResult | None
    ok: bool
    error: str | None = None


@dataclass(frozen=True)
class BatchWorkflowResult:
    output_dir: Path
    package_intake: IntakeFiles
    objects: list[BatchObjectResult]
    report_path: Path

    @property
    def ok(self) -> bool:
        return all(item.ok for item in self.objects)


def run_batch_workflow(
    *,
    pdf_path: Path,
    links_path: Path | None,
    template_path: Path,
    output_dir: Path | None = None,
    profile: str = "apartment",
    complex_name: str | None = None,
    required_count: int | None = None,
    allow_less: bool = False,
    allow_incomplete: bool = False,
    first_page: int | None = None,
    last_page: int | None = None,
    dpi: int = 220,
    force_ocr: bool = False,
    visible: bool = False,
    report_template_path: Path | None = None,
    include_full_screenshots: bool = False,
    progress: ProgressCallback | None = None,
) -> BatchWorkflowResult:
    pdf_file = _resolve_path(pdf_path)
    links_file = _resolve_path(links_path) if links_path else None
    template_file = _resolve_path(template_path)
    out_dir = _resolve_path(output_dir) if output_dir else _default_batch_output_dir(pdf_file)
    out_dir.mkdir(parents=True, exist_ok=True)

    emit_progress(progress, f"Batch workflow: старт пакета. Папка результату: {out_dir}")
    package_intake_dir = out_dir / "00_pdf_intake"
    package_intake = extract_intake_from_pdf(
        pdf_path=pdf_file,
        output_dir=package_intake_dir,
        target_apartment=None,
        template_path=template_file,
        profile=profile,
        complex_name=complex_name,
        links_path=links_file,
        first_page=first_page,
        last_page=last_page,
        dpi=dpi,
        force_ocr=force_ocr,
        progress=progress,
    )
    extracts = package_intake.result.extracts
    if not extracts:
        raise FullWorkflowError("Batch workflow: у PDF не знайдено жодного витягу/об'єкта.")

    emit_progress(progress, f"Batch workflow: знайдено {len(extracts)} об'єкт(ів), формую звіт по кожному.")
    object_results: list[BatchObjectResult] = []
    for index, extract in enumerate(extracts, start=1):
        object_dir = out_dir / _object_dir_name(index, extract)
        result = _run_object_from_package(
            package_intake=package_intake,
            extract=extract,
            object_dir=object_dir,
            links_file=links_file,
            template_file=template_file,
            profile=profile,
            complex_name=complex_name,
            required_count=required_count,
            allow_less=allow_less,
            allow_incomplete=allow_incomplete,
            visible=visible,
            report_template_path=_resolve_path(report_template_path) if report_template_path else None,
            include_full_screenshots=include_full_screenshots,
            progress=progress,
            object_index=index,
            total_objects=len(extracts),
        )
        object_results.append(result)

    report_path = _write_batch_report(out_dir, package_intake=package_intake, object_results=object_results)
    passed = sum(1 for item in object_results if item.ok)
    emit_progress(progress, f"Batch workflow завершено: {passed}/{len(object_results)} звітів без помилок. Зведення: {report_path}")
    return BatchWorkflowResult(
        output_dir=out_dir,
        package_intake=package_intake,
        objects=object_results,
        report_path=report_path,
    )


def _run_object_from_package(
    *,
    package_intake: IntakeFiles,
    extract: ExtractRecord,
    object_dir: Path,
    links_file: Path | None,
    template_file: Path,
    profile: str,
    complex_name: str | None,
    required_count: int | None,
    allow_less: bool,
    allow_incomplete: bool,
    visible: bool,
    report_template_path: Path | None,
    include_full_screenshots: bool,
    progress: ProgressCallback | None,
    object_index: int,
    total_objects: int,
) -> BatchObjectResult:
    apartment = extract.apartment_number
    label = f"[{object_index}/{total_objects}] кв./об'єкт {apartment or 'без номера'} стор. {extract.page}"
    emit_progress(progress, f"Batch workflow {label}: старт.")
    selected_technical = select_technical_passport_for_extract(package_intake.result.technical_passports, extract)
    intake: IntakeFiles | None = None
    excel_workflow: WorkflowResult | None = None
    word_report_path: Path | None = None
    validation: ReportValidationResult | None = None
    try:
        intake = write_intake_selection_files(
            base_result=package_intake.result,
            output_dir=object_dir,
            selected_extract=extract,
            selected_technical=selected_technical,
            template_path=template_file,
            profile=profile,
            complex_name=complex_name,
            links_path=links_file,
        )
        emit_progress(progress, f"Batch workflow {label}: task.generated.yaml створено.")
        excel_workflow = run_excel_workflow(
            task_path=intake.task_yaml,
            links_path=links_file,
            output_dir=object_dir,
            required_count=required_count,
            allow_less=allow_less,
            allow_incomplete=allow_incomplete,
            visible=visible,
            progress=progress,
        )
        if report_template_path:
            word_report_path = object_dir / "valuation_report.docx"
            emit_progress(progress, f"Batch workflow {label}: генерую Word.")
            generate_word_report(
                template_path=report_template_path,
                output_path=word_report_path,
                intake_json=intake.intake_json,
                candidates_json=object_dir / "candidates.json",
                task_path=intake.task_yaml,
                excel_path=excel_workflow.excel.output_path if excel_workflow.excel else None,
                include_full_screenshots=include_full_screenshots,
            )
            emit_progress(progress, f"Batch workflow {label}: validation.")
            validation = validate_report(
                word_path=word_report_path,
                excel_path_value=excel_workflow.excel.output_path if excel_workflow.excel else None,
                intake_json=intake.intake_json,
                task_path=intake.task_yaml,
                candidates_json=object_dir / "candidates.json",
                output_dir=object_dir,
                required_count=required_count or 5,
            )
            if not validation.ok:
                error = f"Validation FAIL: {validation.error_count} errors, {validation.warning_count} warnings"
                emit_progress(progress, f"Batch workflow {label}: {error}.")
                return BatchObjectResult(
                    output_dir=object_dir,
                    extract_page=extract.page,
                    apartment=apartment,
                    intake=intake,
                    excel_workflow=excel_workflow,
                    word_report_path=word_report_path,
                    validation=validation,
                    ok=False,
                    error=error,
                )
        emit_progress(progress, f"Batch workflow {label}: готово.")
        return BatchObjectResult(
            output_dir=object_dir,
            extract_page=extract.page,
            apartment=apartment,
            intake=intake,
            excel_workflow=excel_workflow,
            word_report_path=word_report_path,
            validation=validation,
            ok=True,
        )
    except Exception as exc:
        emit_progress(progress, f"Batch workflow {label}: помилка - {exc}")
        return BatchObjectResult(
            output_dir=object_dir,
            extract_page=extract.page,
            apartment=apartment,
            intake=intake,
            excel_workflow=excel_workflow,
            word_report_path=word_report_path,
            validation=validation,
            ok=False,
            error=str(exc),
        )


def _write_batch_report(
    output_dir: Path,
    *,
    package_intake: IntakeFiles,
    object_results: list[BatchObjectResult],
) -> Path:
    path = output_dir / "batch_report.md"
    passed = sum(1 for item in object_results if item.ok)
    lines = [
        "# Batch Valuation Workflow Report",
        "",
        f"Created: {datetime.now().isoformat(timespec='seconds')}",
        f"Source PDF: {package_intake.result.source_pdf}",
        f"Objects found: {len(package_intake.result.extracts)}",
        f"Reports passed: {passed}",
        f"Reports failed: {len(object_results) - passed}",
        f"Package intake: {package_intake.intake_json}",
        "",
        "## Objects",
        "",
        "| # | Apartment/Object | Extract page | Status | Folder | Word | Validation | Error |",
        "| ---: | --- | ---: | --- | --- | --- | --- | --- |",
    ]
    for index, item in enumerate(object_results, start=1):
        validation_path = item.validation.validation_md if item.validation else ""
        lines.append(
            "| "
            f"{index} | "
            f"{item.apartment or ''} | "
            f"{item.extract_page if item.extract_page is not None else ''} | "
            f"{'PASS' if item.ok else 'FAIL'} | "
            f"{item.output_dir} | "
            f"{item.word_report_path or ''} | "
            f"{validation_path} | "
            f"{_table_escape(item.error or '')} |"
        )
    path.write_text("\n".join(lines).strip() + "\n", encoding="utf-8")
    return path


def _object_dir_name(index: int, extract: ExtractRecord) -> str:
    if extract.apartment_number:
        slug = f"apt_{extract.apartment_number}"
    elif extract.registry_object_number:
        slug = f"registry_{extract.registry_object_number}"
    else:
        slug = f"page_{extract.page}"
    slug = re.sub(r"[^A-Za-z0-9_-]+", "-", slug).strip("-").lower()
    return f"{index:02d}_{slug or 'object'}"


def _table_escape(value: str) -> str:
    return value.replace("|", "\\|").replace("\n", " ")


def _resolve_path(path: Path) -> Path:
    return path if path.is_absolute() else PROJECT_ROOT / path


def _default_batch_output_dir(pdf_path: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    slug = re.sub(r"[^A-Za-z0-9_-]+", "-", pdf_path.stem).strip("-").lower()
    if not slug:
        slug = "batch"
    return ensure_output_dir(f"{timestamp}_batch_{slug[:40]}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run batch PDF intake -> one valuation report per real-estate object.")
    parser.add_argument("--pdf", type=Path, required=True, help="Input extract/technical-passport PDF.")
    parser.add_argument("--links", type=Path, default=None, help="Optional UTF-8 text file with listing URLs.")
    parser.add_argument("--template", type=Path, required=True, help="Excel .xls template path.")
    parser.add_argument("--out", type=Path, default=None, help="Output directory.")
    parser.add_argument("--profile", default="apartment", help="Template profile name.")
    parser.add_argument("--complex-name", default=None)
    parser.add_argument("--required-count", type=int, default=None)
    parser.add_argument("--allow-less", action="store_true")
    parser.add_argument("--allow-incomplete", action="store_true")
    parser.add_argument("--first-page", type=int, default=None)
    parser.add_argument("--last-page", type=int, default=None)
    parser.add_argument("--dpi", type=int, default=220)
    parser.add_argument("--force-ocr", action="store_true")
    parser.add_argument("--visible", action="store_true")
    parser.add_argument("--report-template", type=Path, default=None)
    parser.add_argument("--include-full-screenshots", action="store_true")
    args = parser.parse_args(argv)

    console = Console()
    try:
        result = run_batch_workflow(
            pdf_path=args.pdf,
            links_path=args.links,
            template_path=args.template,
            output_dir=args.out,
            profile=args.profile,
            complex_name=args.complex_name,
            required_count=args.required_count,
            allow_less=args.allow_less,
            allow_incomplete=args.allow_incomplete,
            first_page=args.first_page,
            last_page=args.last_page,
            dpi=args.dpi,
            force_ocr=args.force_ocr,
            visible=args.visible,
            report_template_path=args.report_template,
            include_full_screenshots=args.include_full_screenshots,
        )
    except Exception as exc:
        console.print(f"[red]Batch workflow failed:[/red] {exc}")
        return 1

    passed = sum(1 for item in result.objects if item.ok)
    console.print("[green]Batch workflow complete[/green]" if result.ok else "[yellow]Batch workflow completed with failures[/yellow]")
    console.print(f"Reports passed: {passed}/{len(result.objects)}")
    console.print(f"Report: {result.report_path}")
    console.print(f"Output: {result.output_dir}")
    return 0 if result.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
