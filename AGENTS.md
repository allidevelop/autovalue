# Project Instructions

- This repository contains the Realtify/Autovalue real-estate valuation MVP.
- Keep client PDFs, source Word/Excel files, generated reports, screenshots, builds, virtual environments, and browser/OCR binaries out of Git.
- The local Windows app is a wrapper over the Python engine in `src/realtify`.
- Batch mode is the default product workflow: one uploaded PDF can contain multiple extracts, and the system must generate one report folder per detected real-estate object.
- Use `scripts/run_batch_workflow.py` for end-to-end package runs and `scripts/run_full_workflow.py` only for single-object debugging.
- Do not rely on LibreOffice on this workstation for DOCX QA; it is known to have a broken local configuration. Use structural DOCX validation unless LibreOffice is repaired.
- Build the Windows package with `scripts/build_windows_app.ps1`; do not commit `dist/` or `build/`.
