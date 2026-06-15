# realtify-tests

Local MVP workspace for automating real estate valuation workflows.

## Current scope

- Parse incoming client documents.
- OCR scanned PDFs when no text layer is available.
- Collect comparable listings and full-page screenshots.
- Fill Excel valuation templates.
- Generate final valuation reports from Word templates.

## Client workflow decisions

- Final deliverable is Word (`.docx` preferred; `.doc` acceptable). PDF export is optional and can be done manually.
- Existing `.doc` reports are examples/templates and should be converted into `.docx` templates with placeholders.
- Incoming client documents may arrive as separate files or one bundle. In a bundle, the expected order is extract first, then technical passport.
- Apartment room count comes from the technical passport explication: count rows where column 6 contains living area.
- Bargaining adjustment is not automated in the MVP and must not be overwritten.
- Store full-page listing screenshots for audit, but insert readable compressed images into the Word report.

See [docs/REPORT_VARIABLES.md](docs/REPORT_VARIABLES.md) for the first report placeholder map.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m playwright install chromium
.\scripts\bootstrap_tools.ps1
.\.venv\Scripts\python.exe -m pip install -e .
.\.venv\Scripts\python.exe scripts\check_env.py
```

Tesseract can use bundled language data from `tools/tessdata` or the system tessdata directory, for example `/usr/share/tesseract-ocr/5/tessdata` on Linux.
Portable Poppler is stored in `tools/poppler`.

To smoke-test OCR on a scanned PDF:

```powershell
.\.venv\Scripts\python.exe scripts\smoke_ocr_pdf.py "C:\path\to\extract.pdf"
```

To extract a draft task file from a scanned extract/technical-passport PDF:

```powershell
.\.venv\Scripts\python.exe scripts\extract_intake.py `
  --pdf "C:\path\to\bundle.pdf" `
  --apartment 353 `
  --template "C:\path\to\template.xls" `
  --profile apartment
```

The output folder contains `intake.json`, per-page OCR text, rendered page images, and `task.generated.yaml`. OCR-derived fields should be reviewed when warnings are present, especially room count from scanned explication tables.

To collect listing candidates from direct links:

```powershell
.\.venv\Scripts\python.exe scripts\collect_from_links.py `
  --links inputs\links.example.txt `
  --property-type apartment `
  --transaction-type sale
```

The output folder contains `candidates.json`, `report.md`, archive screenshots under `screenshots/`, and compressed Word-ready images under `report_images/`. HTTP error pages, deleted listings, captcha pages, and blocked pages are written to the `errors` section with their diagnostic screenshot path and are not counted as candidates.

Configured sources live in [config/sources.yaml](config/sources.yaml). Current primary sources are `rieltor`, `dimria`, `lun`, and `real_estate_lviv`; `olx` is present but disabled by default. Rieltor agent subdomains such as `0988666715.rieltor.ua` are detected as Rieltor links.

To discover listing links automatically from configured source catalogs:

```powershell
.\.venv\Scripts\python.exe scripts\discover_links.py `
  --task config\task.example.yaml `
  --max-links 25
```

Discovery writes `discovered_links.txt` and `discovery.json`. For apartments with `collection.only_newbuilds: true`, Rieltor discovery uses the `/newhouse/` catalog. Direct `--links` files still work and override automatic discovery.

After collection, the workflow ranks candidates before Excel/Word generation. The full parsed pool is saved as `collected_candidates.json`; the final selected analogs are saved as `candidates.json`; selection scores and rejection reasons are saved as `candidate_selection.json` and `candidate_selection.md`.

To fill an Excel valuation template from collected candidates:

```powershell
.\.venv\Scripts\python.exe scripts\fill_excel_template.py `
  --task config\task.example.yaml `
  --candidates outputs\<run_folder>\candidates.json
```

By default the script requires 5 complete analogs. It creates a copy of the source `.xls` in `outputs/<timestamp>_excel/` and writes only cells declared in `config/template_profiles/*.yaml`. The default backend is `auto`: it uses Microsoft Excel COM when available and falls back to the cross-platform Python `.xls` backend otherwise. To force the cross-platform backend:

```powershell
$env:REALTIFY_EXCEL_ENGINE='python-xls'
```

The Python backend writes a sidecar file next to the workbook: `*.xls.realtify.json`. Word generation and validation read calculation summaries and adjustment rows from that sidecar first, so server runs do not need Microsoft Excel.

To run the full links-to-Excel workflow:

```powershell
.\.venv\Scripts\python.exe scripts\run_excel_workflow.py `
  --task config\task.example.yaml
```

The workflow writes one result folder with `collected_candidates.json`, selected `candidates.json`, `candidate_selection.json`, `report.md`, full-page screenshots, compressed report images, and the filled `.xls` workbook. It fails by default if fewer than 5 suitable analogs are selected; use `--allow-less` only for smoke tests.

To run the batch PDF-to-report workflow, producing one report folder per object found in the PDF:

```powershell
.\.venv\Scripts\python.exe scripts\run_batch_workflow.py `
  --pdf "C:\path\to\bundle.pdf" `
  --template "C:\path\to\template.xls" `
  --profile apartment `
  --report-template config\report_templates\valuation_report_real_template.docx
```

The batch workflow reads/OCRs the PDF once, then creates `00_pdf_intake/` plus one numbered object folder per extract, for example `01_apt_353/`. Each object folder contains its own `intake.json`, `task.generated.yaml`, analog collection, Excel workbook, Word report, and validation files. The root folder contains `batch_report.md` with PASS/FAIL status and links for every object.

To run the legacy single-object PDF-to-Excel workflow in one command:

```powershell
.\.venv\Scripts\python.exe scripts\run_full_workflow.py `
  --pdf "C:\path\to\bundle.pdf" `
  --apartment 353 `
  --template "C:\path\to\template.xls" `
  --profile apartment
```

The single-object workflow creates one result folder containing OCR output, `intake.json`, `task.generated.yaml`, `discovery.json`, `discovered_links.txt`, collected listing data, selected analogs, screenshots, report-ready images, `report.md`, and the filled Excel workbook. Pass `--links inputs\links.txt` only when you want to force a reviewed manual link list.

To prepare a Word report template:

```powershell
.\.venv\Scripts\python.exe scripts\prepare_report_template.py `
  --source "C:\path\to\old-report.docx" `
  --out outputs\report_template
```

The report generator requires a `.docx` template with placeholders like `{{address_full}}`, `{{extract_index_number}}`, `{{comparables_table}}`, and `{{report_listing_images}}`. Old `.doc` files can be passed to the prepare command, but if local Word cannot repair/convert the file, open it manually and save it as `.docx`.

To create the real valuation report template from a highlighted report sample:

```powershell
.\.venv\Scripts\python.exe scripts\create_real_report_template.py `
  --source "C:\path\to\converted-report.docx" `
  --out config\report_templates\valuation_report_real_template.docx
```

The command replaces recognized highlighted values with placeholders and writes `valuation_report_real_template.inventory.md` next to the template. Highlights listed under manual review are intentionally left unchanged until their source data is formalized.

To create a minimal fallback template for smoke tests:

```powershell
.\.venv\Scripts\python.exe scripts\prepare_report_template.py `
  --create-default config\report_templates\default_report_template.docx
```

To generate only the Word report from an existing workflow output:

```powershell
.\.venv\Scripts\python.exe scripts\generate_word_report.py `
  --template config\report_templates\default_report_template.docx `
  --intake outputs\<run_folder>\intake.json `
  --task outputs\<run_folder>\task.generated.yaml `
  --candidates outputs\<run_folder>\candidates.json `
  --excel outputs\<run_folder>\apartment_filled.xls `
  --out outputs\<run_folder>\valuation_report.docx
```

To validate a generated Word report against the selected object, analogs, and Excel calculation:

```powershell
.\.venv\Scripts\python.exe scripts\validate_report.py `
  --word outputs\<run_folder>\valuation_report.docx `
  --excel outputs\<run_folder>\apartment_filled.xls `
  --intake outputs\<run_folder>\intake.json `
  --task outputs\<run_folder>\task.generated.yaml `
  --candidates outputs\<run_folder>\candidates.json `
  --out outputs\<run_folder>
```

Validation writes `validation.json` and `validation.md`. It checks unresolved placeholders, target apartment/area drift, selected analog count, active listing links, screenshot files, the Word comparables table, and the Word adjustment table against the filled Excel workbook. The full workflow runs this validation automatically when a Word template is provided.

To include Word generation in the full workflow:

```powershell
.\.venv\Scripts\python.exe scripts\run_full_workflow.py `
  --pdf "C:\path\to\bundle.pdf" `
  --apartment 353 `
  --template "C:\path\to\template.xls" `
  --profile apartment `
  --report-template config\report_templates\valuation_report_real_template.docx
```

To launch the local Windows UI wrapper:

```powershell
.\.venv\Scripts\python.exe scripts\run_windows_app.py
```

or:

```powershell
.\.venv\Scripts\realtify-windows-app.exe
```

The UI is a thin wrapper over the same workflow. By default it runs batch mode and generates a separate report for every object found in the PDF; the single-object mode is still available by unchecking the batch option and choosing a target object number. See [docs/WINDOWS_APP_MVP.md](docs/WINDOWS_APP_MVP.md).

For manual smoke testing through the UI, use `client_test_files/`: it contains the apartment 353 PDF, Excel template, Word template, a ready output folder, and a short test-run checklist.

The web-version direction is documented in [docs/WEB_ARCHITECTURE_MVP.md](docs/WEB_ARCHITECTURE_MVP.md).

To launch the FastAPI web preview:

```powershell
.\.venv\Scripts\python.exe scripts\run_web_app.py --host 127.0.0.1 --port 8765
```

Then open `http://127.0.0.1:8765`. The web UI accepts a PDF bundle, Excel template, optional Word template, optional link list, object type, and page range. It runs the batch workflow in a background thread, shows timestamped progress events, and publishes a review zip when the job finishes.

Set `REALTIFY_WEB_USERNAME` and `REALTIFY_WEB_PASSWORD` to enable the built-in login page and protect every web route with a signed session cookie:

```powershell
$env:REALTIFY_WEB_USERNAME = "autovalue"
$env:REALTIFY_WEB_PASSWORD = "change-this-password"
```

The web app can now run report generation on Linux when the uploaded Excel template is legacy `.xls`. Microsoft Excel COM remains supported on Windows, but is no longer required for the server path.

Valuation date source is resolved in this order:

1. `task.valuation_date` or `target.valuation_date`.
2. Optional Excel source config, for example `valuation_date_source: {type: excel_cell, path: template, sheet: Sheet1, cell: B2}`.
3. Current date.

To build the Windows `.exe` package:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_windows_app.ps1
```

The build output is `dist\Realtify\Realtify.exe`. The build script bundles `config`, Poppler, OCR language data, and Playwright Chromium under the PyInstaller `onedir` output. Microsoft Excel and Tesseract OCR still must be installed on the appraiser's Windows machine.

## Notes

Do not overwrite client source files or template files. Generated files must go under `outputs/<task_id>/`.
