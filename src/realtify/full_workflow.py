from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from rich.console import Console

from realtify.intake import IntakeFiles, extract_intake_from_pdf
from realtify.paths import PROJECT_ROOT, ensure_output_dir
from realtify.progress import ProgressCallback, emit_progress
from realtify.report_generator import generate_word_report
from realtify.report_validator import ReportValidationResult, summarize_validation_failure, validate_report
from realtify.workflow import WorkflowResult, run_excel_workflow


class FullWorkflowError(RuntimeError):
    pass


@dataclass(frozen=True)
class FullWorkflowResult:
    output_dir: Path
    intake: IntakeFiles
    excel_workflow: WorkflowResult
    report_path: Path
    word_report_path: Path | None = None
    validation: ReportValidationResult | None = None


def run_full_workflow(
    *,
    pdf_path: Path,
    links_path: Path | None,
    template_path: Path,
    output_dir: Path | None = None,
    apartment: str | None = None,
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
    report_output_path: Path | None = None,
    include_full_screenshots: bool = False,
    progress: ProgressCallback | None = None,
) -> FullWorkflowResult:
    pdf_file = _resolve_path(pdf_path)
    links_file = _resolve_path(links_path) if links_path else None
    template_file = _resolve_path(template_path)
    out_dir = _resolve_path(output_dir) if output_dir else _default_full_output_dir(pdf_file, apartment)
    out_dir.mkdir(parents=True, exist_ok=True)

    emit_progress(progress, f"Старт workflow. Папка результату: {out_dir}")
    emit_progress(progress, "PDF intake/OCR: читаю витяг та техпаспорт.")
    intake = extract_intake_from_pdf(
        pdf_path=pdf_file,
        output_dir=out_dir,
        target_apartment=apartment,
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
    selected = intake.result.selected_extract
    selected_tp = intake.result.selected_technical_passport
    emit_progress(
        progress,
        "PDF intake/OCR завершено: "
        f"сторінка витягу {selected.page if selected else 'не знайдено'}, "
        f"квартира {selected.apartment_number if selected else 'не знайдено'}, "
        f"площа {selected.total_area_m2 if selected and selected.total_area_m2 is not None else 'не знайдено'}, "
        f"техпаспорт {selected_tp.apartment_number if selected_tp else 'не знайдено'} "
        f"(сторінки {', '.join(str(page) for page in selected_tp.pages) if selected_tp else 'не знайдено'}).",
    )
    emit_progress(progress, "Переходжу до пошуку/збору аналогів та Excel.")
    excel_workflow = run_excel_workflow(
        task_path=intake.task_yaml,
        links_path=links_file,
        output_dir=out_dir,
        required_count=required_count,
        allow_less=allow_less,
        allow_incomplete=allow_incomplete,
        visible=visible,
        progress=progress,
    )

    word_report_path: Path | None = None
    validation: ReportValidationResult | None = None
    if report_template_path:
        report_template = _resolve_path(report_template_path)
        word_report_path = _resolve_path(report_output_path) if report_output_path else out_dir / "valuation_report.docx"
        emit_progress(progress, "Word-звіт: генерую фінальний DOCX за шаблоном.")
        generate_word_report(
            template_path=report_template,
            output_path=word_report_path,
            intake_json=intake.intake_json,
            candidates_json=out_dir / "candidates.json",
            task_path=intake.task_yaml,
            excel_path=excel_workflow.excel.output_path if excel_workflow.excel else None,
            include_full_screenshots=include_full_screenshots,
        )
        emit_progress(progress, f"Word-звіт створено: {word_report_path}")
        emit_progress(progress, "Validation: перевіряю Word проти intake, аналогів та Excel.")
        validation = validate_report(
            word_path=word_report_path,
            excel_path_value=excel_workflow.excel.output_path if excel_workflow.excel else None,
            intake_json=intake.intake_json,
            task_path=intake.task_yaml,
            candidates_json=out_dir / "candidates.json",
            output_dir=out_dir,
            required_count=required_count or 5,
        )
        emit_progress(progress, f"Validation завершено: {'PASS' if validation.ok else 'FAIL'} ({validation.error_count} errors, {validation.warning_count} warnings).")

    report_path = _write_full_report(
        out_dir,
        intake=intake,
        excel_workflow=excel_workflow,
        word_report_path=word_report_path,
        validation=validation,
    )
    if validation and not validation.ok:
        raise FullWorkflowError(summarize_validation_failure(validation))
    emit_progress(progress, "Workflow повністю завершено.")
    return FullWorkflowResult(
        output_dir=out_dir,
        intake=intake,
        excel_workflow=excel_workflow,
        report_path=report_path,
        word_report_path=word_report_path,
        validation=validation,
    )


def _write_full_report(
    output_dir: Path,
    *,
    intake: IntakeFiles,
    excel_workflow: WorkflowResult,
    word_report_path: Path | None = None,
    validation: ReportValidationResult | None = None,
) -> Path:
    report_path = output_dir / "report.md"
    selected_extract = intake.result.selected_extract
    selected_technical = intake.result.selected_technical_passport
    excel_result = excel_workflow.excel

    lines = [
        "# Full Valuation Workflow Report",
        "",
        f"Created: {datetime.now().isoformat(timespec='seconds')}",
        f"Source PDF: {intake.result.source_pdf}",
        f"Intake JSON: {intake.intake_json}",
        f"Intake summary: {output_dir / 'intake_summary.md'}",
        f"Generated task YAML: {intake.task_yaml}",
        f"Candidates JSON: {output_dir / 'candidates.json'}",
        f"Collected candidates JSON: {output_dir / 'collected_candidates.json'}",
        f"Candidate selection JSON: {output_dir / 'candidate_selection.json'}",
        f"Excel output: {excel_result.output_path if excel_result else 'not created'}",
        f"Word output: {word_report_path if word_report_path else 'not created'}",
        f"Validation JSON: {validation.validation_json if validation else 'not created'}",
        f"Validation report: {validation.validation_md if validation else 'not created'}",
        f"Validation status: {'PASS' if validation and validation.ok else ('FAIL' if validation else 'not run')}",
        f"Full-page screenshots: {output_dir / 'screenshots'}",
        f"Report images: {output_dir / 'report_images'}",
        "",
        "## Selected Object",
        "",
    ]
    if selected_extract:
        lines.extend(
            [
                f"- Apartment: {selected_extract.apartment_number or 'not found'}",
                f"- Extract page: {selected_extract.page}",
                f"- Address: {selected_extract.address_full or 'not found'}",
                f"- Extract index number: {selected_extract.extract_index_number or 'not found'}",
                f"- Extract formed at: {selected_extract.extract_formed_at or 'not found'}",
                f"- Object identifier: {selected_extract.object_identifier or 'not found'}",
                f"- Total area: {selected_extract.total_area_m2 if selected_extract.total_area_m2 is not None else 'not found'}",
                f"- Living area: {selected_extract.living_area_m2 if selected_extract.living_area_m2 is not None else 'not found'}",
            ]
        )
    else:
        lines.append("- Extract: not found")

    lines.extend(["", "## Technical Passport", ""])
    if selected_technical:
        lines.extend(
            [
                f"- Apartment: {selected_technical.apartment_number or 'not found'}",
                f"- Registration number: {selected_technical.registration_number or 'not found'}",
                f"- Technical passport date: {selected_technical.technical_passport_date or 'not found'}",
                f"- Floor: {selected_technical.floor_or_level or 'not found'}",
                f"- Rooms: {selected_technical.rooms_count if selected_technical.rooms_count is not None else 'not found'}",
                f"- Pages: {', '.join(str(page) for page in selected_technical.pages) or 'not found'}",
                f"- Warnings: {', '.join(selected_technical.warnings) if selected_technical.warnings else 'none'}",
            ]
        )
    else:
        lines.append("- Technical passport: not found")

    lines.extend(
        [
            "",
            "## Comparables",
            "",
            f"- Collected: {len(excel_workflow.raw_collection.candidates) if excel_workflow.raw_collection else len(excel_workflow.collection.candidates)}",
            f"- Selected for report: {len(excel_workflow.collection.candidates)}",
            f"- Rejected links/errors: {len(excel_workflow.collection.errors)}",
        ]
    )
    if excel_workflow.selection and excel_workflow.selection.warnings:
        lines.append(f"- Selection warnings: {', '.join(excel_workflow.selection.warnings)}")
    if excel_result and excel_result.warnings:
        lines.append(f"- Excel warnings: {', '.join(excel_result.warnings)}")

    lines.extend(["", "## Candidate Details", ""])
    for idx, candidate in enumerate(excel_workflow.collection.candidates, start=1):
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

    if excel_workflow.collection.errors:
        lines.extend(["## Rejected Links", ""])
        for error in excel_workflow.collection.errors:
            extra = f" screenshot={error['screenshot_path']}" if error.get("screenshot_path") else ""
            lines.append(f"- `{error['source']}` {error['url']}: {error['error']}{extra}")

    if validation and validation.issues:
        lines.extend(["", "## Validation Issues", ""])
        for issue in validation.issues:
            location = f" ({issue.location})" if issue.location else ""
            lines.append(f"- {issue.severity.upper()} `{issue.code}`{location}: {issue.message}")

    if intake.result.warnings:
        lines.extend(["", "## Intake Warnings", ""])
        lines.extend(f"- {warning}" for warning in intake.result.warnings)

    report_path.write_text("\n".join(lines).strip() + "\n", encoding="utf-8")
    return report_path


def _resolve_path(path: Path) -> Path:
    return path if path.is_absolute() else PROJECT_ROOT / path


def _default_full_output_dir(pdf_path: Path, apartment: str | None) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = pdf_path.stem
    suffix = f"apt_{apartment}" if apartment else stem
    slug = re.sub(r"[^A-Za-z0-9_-]+", "-", suffix).strip("-").lower()
    if not slug:
        slug = "full_workflow"
    return ensure_output_dir(f"{timestamp}_{slug[:40]}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run PDF intake -> listing collection -> filled Excel workflow.")
    parser.add_argument("--pdf", type=Path, required=True, help="Input extract/technical-passport PDF.")
    parser.add_argument("--links", type=Path, default=None, help="Optional UTF-8 text file with listing URLs. If omitted, sources are discovered automatically.")
    parser.add_argument("--template", type=Path, required=True, help="Excel .xls template path.")
    parser.add_argument("--out", type=Path, default=None, help="Output directory.")
    parser.add_argument("--apartment", default=None, help="Target apartment number in bundled PDFs.")
    parser.add_argument("--profile", default="apartment", help="Template profile name.")
    parser.add_argument("--complex-name", default=None)
    parser.add_argument("--required-count", type=int, default=None)
    parser.add_argument("--allow-less", action="store_true", help="Allow fewer candidates than required.")
    parser.add_argument("--allow-incomplete", action="store_true", help="Allow missing critical candidate fields.")
    parser.add_argument("--first-page", type=int, default=None)
    parser.add_argument("--last-page", type=int, default=None)
    parser.add_argument("--dpi", type=int, default=220)
    parser.add_argument("--force-ocr", action="store_true")
    parser.add_argument("--visible", action="store_true", help="Show Excel while filling.")
    parser.add_argument("--report-template", type=Path, default=None, help="Prepared .docx report template.")
    parser.add_argument("--report-out", type=Path, default=None, help="Output .docx report path.")
    parser.add_argument("--include-full-screenshots", action="store_true", help="Insert full-page screenshots into Word appendix.")
    args = parser.parse_args(argv)

    console = Console()
    try:
        result = run_full_workflow(
            pdf_path=args.pdf,
            links_path=args.links,
            template_path=args.template,
            output_dir=args.out,
            apartment=args.apartment,
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
            report_output_path=args.report_out,
            include_full_screenshots=args.include_full_screenshots,
        )
    except Exception as exc:
        console.print(f"[red]Full workflow failed:[/red] {exc}")
        return 1

    selected = result.intake.result.selected_extract
    console.print("[green]Full workflow complete[/green]")
    console.print(f"Selected apartment: {selected.apartment_number if selected else 'not found'}")
    raw_count = len(result.excel_workflow.raw_collection.candidates) if result.excel_workflow.raw_collection else len(result.excel_workflow.collection.candidates)
    console.print(f"Candidates collected: {raw_count}")
    console.print(f"Candidates selected: {len(result.excel_workflow.collection.candidates)}")
    console.print(f"Errors: {len(result.excel_workflow.collection.errors)}")
    console.print(f"Excel: {result.excel_workflow.excel.output_path if result.excel_workflow.excel else 'not created'}")
    console.print(f"Word: {result.word_report_path if result.word_report_path else 'not created'}")
    console.print(f"Validation: {'PASS' if result.validation and result.validation.ok else ('not run' if not result.validation else 'FAIL')}")
    if result.validation:
        console.print(f"Validation report: {result.validation.validation_md}")
    console.print(f"Report: {result.report_path}")
    console.print(f"Output: {result.output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
