# Windows App MVP

## Goal

Wrap the existing local valuation workflow in a Windows-friendly UI for the appraiser:

1. Select the incoming PDF bundle or extract.
2. Enter the target apartment/object number.
3. Select Excel and Word templates.
4. Run automatic comparable discovery, screenshots, selection, Excel filling, and Word generation.
5. Open the result folder.

The UI must stay thin. The valuation logic remains in the CLI/workflow modules so scripts and app produce the same files.

## Current MVP

Entry points:

- `scripts/run_windows_app.py`
- `realtify-windows-app`

The app uses Python `tkinter`, which ships with standard Python and avoids adding GUI dependencies at this stage.

Current controls:

- PDF file picker.
- Target apartment/object number.
- Property type profile.
- Excel template picker.
- Word template picker.
- Output folder picker.
- Optional complex name.
- Full-page screenshot appendix toggle.
- Visible Excel toggle.
- Environment check button.
- Generate report button.
- Open result folder button.
- Open generated Word report button.
- Open generated Excel workbook button.
- Open validation report button.
- Visible validation status in the status bar.

## Output Contract

The app writes the same output folder as `scripts/run_full_workflow.py`:

- `intake.json`
- `task.generated.yaml`
- `discovery.json`
- `discovered_links.txt`
- `collected_candidates.json`
- `candidate_selection.json`
- `candidate_selection.md`
- `candidates.json`
- `apartment_filled.xls`
- `valuation_report.docx`
- `validation.json`
- `validation.md`
- `runtime.log`
- `screenshots/*.png`
- `report_images/*.jpg`

## Packaging Direction

Draft build script:

```powershell
.\scripts\build_windows_app.ps1
```

Default build output:

```text
dist\Realtify\Realtify.exe
```

The build script installs Playwright Chromium into `tools\ms-playwright` unless `-SkipBrowserInstall` is passed, then bundles `config`, Poppler, OCR language data, and Playwright browsers into the PyInstaller `onedir` output.

Future packaging work:

- Add a formal installer or zip release script.
- Decide whether Excel/Word templates are bundled or selected from an external folder.
- Create a desktop shortcut.
- Add a persistent app settings file for default template paths.
- Add a review screen for extracted PDF data before report generation.
- Add a review/edit screen for selected analogs before Excel/Word generation.

## Constraints

- Microsoft Excel must be installed on the user's machine because `.xls` templates are filled through Excel COM.
- Tesseract OCR must be installed on the user's machine; the app bundles OCR language files, not the Tesseract executable.
- LibreOffice is not required for report generation.
- Visual DOCX render QA is currently not part of the app because the local LibreOffice installation is unreliable.
