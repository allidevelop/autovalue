# Realtify Web Architecture MVP

## Goal

Move the current local Windows workflow into a browser-based product:

1. Appraiser opens a web page.
2. Uploads PDF bundle and optional templates.
3. Enters target object number and property type.
4. Starts a background valuation job.
5. Watches live progress logs.
6. Downloads a package of reports: one object folder per real-estate object found in the PDF.

The web version should reuse the existing Python engine modules first. UI and deployment can change; valuation logic should not be rewritten until the current workflow is stable.

## Current Engine To Reuse

- `realtify.intake`: PDF/OCR intake.
- `realtify.discover_links`: source catalog discovery.
- `realtify.collect_from_links`: listing screenshots and extraction.
- `realtify.candidate_selector`: analog ranking.
- `realtify.fill_template`: Excel output.
- `realtify.report_generator`: Word report.
- `realtify.report_validator`: structural validation.
- `realtify.full_workflow`: single-object orchestration.
- `realtify.batch_workflow`: package orchestration, one report per object found in the PDF.

The new progress callback can become the web job event stream.

## Recommended MVP Stack

- Backend: `FastAPI`
- Worker: `RQ + Redis` for MVP, `Celery` later if needed.
- Frontend: `Next.js` or `React + Vite`.
- Database: `PostgreSQL`.
- File storage MVP: local mounted volume.
- File storage production: S3-compatible storage.
- Runtime: Docker Compose.

## Core Services

### API Service

Responsible for:

- Authentication later.
- File upload.
- Job creation.
- Job status.
- Artifact download.
- Streaming or polling progress logs.

### Worker Service

Responsible for:

- Running `run_full_workflow(...)`.
- Writing artifacts into a job folder.
- Writing structured job events.
- Marking job status: `queued`, `running`, `passed`, `failed`.

### Redis

Responsible for:

- Job queue.
- Lightweight event/pubsub if we use live logs.

### PostgreSQL

Tables:

- `users`
- `valuation_jobs`
- `job_artifacts`
- `job_events`
- `templates`

## Minimal API Contract

### Create Job

`POST /api/jobs`

Multipart fields:

- `pdf_file`
- `apartment_number`
- `property_type`
- `excel_template_id` or `excel_template_file`
- `word_template_id` or `word_template_file`
- `complex_name`
- `include_full_screenshots`
- `batch_all_objects` (default `true`)

Response:

```json
{
  "job_id": "uuid",
  "status": "queued"
}
```

### Job Status

`GET /api/jobs/{job_id}`

Response:

```json
{
  "job_id": "uuid",
  "status": "running",
  "progress_message": "Збір оголошень: [14/20] ...",
  "created_at": "...",
  "started_at": "...",
  "finished_at": null,
  "validation_ok": null
}
```

### Job Events

`GET /api/jobs/{job_id}/events`

Returns ordered log events. MVP can poll every 2 seconds. Later: Server-Sent Events.

### Artifacts

`GET /api/jobs/{job_id}/artifacts`

Returns:

- `valuation_report.docx`
- `apartment_filled.xls` or future `.xlsx`
- `validation.md`
- `validation.json`
- `candidates.json`
- `candidate_selection.md`
- screenshot archive

## Job Folder Contract

For MVP, keep a folder per job:

```text
storage/jobs/<job_id>/
  input/
    source.pdf
    excel_template.xls
    word_template.docx
  output/
    batch_report.md
    runtime.log
    00_pdf_intake/
      intake.json
      intake_summary.md
    01_apt_353/
      intake.json
      task.generated.yaml
      discovered_links.txt
      discovery.json
      collected_candidates.json
      candidates.json
      apartment_filled.xls
      valuation_report.docx
      validation.md
      validation.json
      screenshots/
      report_images/
    02_apt_373/
      ...
```

## Main Technical Blocker

The current Excel writer uses Microsoft Excel COM, which is Windows-only.

Migration options:

1. Short-term web MVP on a Windows server with Excel installed.
   - Fastest, but fragile and not cloud-native.
2. Convert `.xls` templates to `.xlsx` and write with `openpyxl`.
   - Best long-term direction.
3. Use LibreOffice headless in Docker for recalculation.
   - Cross-platform, but must validate formatting/calculation accuracy.
4. Move valuation math into Python and generate Excel only as an output artifact.
   - Most robust long-term, but requires formalizing all formulas.

Recommended path:

Start web architecture now, but keep production Windows `.exe` until Excel COM is removed or isolated.

## Web MVP Phases

### Phase 1: API Wrapper Around Existing Engine

- Create FastAPI backend.
- Add `/jobs` create/status/events/artifacts endpoints.
- Run jobs synchronously or with a simple background worker.
- Store uploaded files and artifacts locally.
- Reuse existing `run_full_workflow(progress=...)`.

### Phase 2: Frontend Dashboard

- Upload form.
- Object parameters.
- Live progress log.
- Result card with validation PASS/FAIL.
- Download buttons.

### Phase 3: Worker Queue

- Add Redis + RQ.
- Make long report generation independent from API requests.
- Add cancel/retry.

### Phase 4: Template Management

- Upload and save Excel/Word templates.
- Set default templates.
- Version templates.

### Phase 5: Cross-Platform Excel Refactor

- Convert templates to `.xlsx`.
- Replace Excel COM with a cross-platform engine.
- Add regression checks comparing old Excel output against new output.

## UX Requirements

- Always show current stage and timestamped event log.
- Never leave the user staring at a static screen during long collection.
- Keep final state explicit:
  - `Validation PASS`: green.
  - `Validation FAIL`: red with issue list.
  - `Needs manual review`: yellow.
- Let user download all artifacts as one zip.

## Deployment Options

### Local Docker

Good for one appraiser or office machine.

Pros:

- Data stays local.
- Easier privacy story.

Cons:

- Docker setup burden.
- Excel COM cannot run in Linux containers.

### Cloud VPS

Good for multi-user access.

Pros:

- One central system.
- Easy browser access from any device.

Cons:

- Must solve Excel COM.
- Need authentication, backups, and file privacy.

### Hybrid

Use Windows `.exe` now for production tests, develop web MVP in parallel, then migrate once formulas and templates are stable.

This is the recommended project path.
