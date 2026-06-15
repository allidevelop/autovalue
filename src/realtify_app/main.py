from __future__ import annotations

import os
import queue
import threading
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from tkinter import BooleanVar, StringVar, Text, Tk, filedialog, messagebox
from tkinter import font as tkfont
from tkinter import ttk

from realtify.batch_workflow import run_batch_workflow
from realtify.env_check import collect_status
from realtify.full_workflow import run_full_workflow
from realtify.intake import extract_intake_from_pdf
from realtify.paths import PROJECT_ROOT, RESOURCE_ROOT


APP_TITLE = "Realtify - формування звіту оцінки"

PALETTE = {
    "app_bg": "#eef3f8",
    "card_bg": "#ffffff",
    "border": "#c9d6e2",
    "text": "#17202a",
    "muted": "#5d6b7a",
    "primary": "#1f6feb",
    "primary_dark": "#174ea6",
    "secondary": "#e8eef6",
    "success": "#188038",
    "warning": "#b7791f",
    "log_bg": "#0f172a",
    "log_fg": "#dbeafe",
}


def main() -> int:
    app = RealtifyApp()
    app.mainloop()
    return 0


class RealtifyApp(Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1220x820")
        self.minsize(1120, 760)
        self.configure(bg=PALETTE["app_bg"])
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.messages: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.analysis_worker: threading.Thread | None = None
        self.last_output_dir: Path | None = None
        self.last_excel_path: Path | None = None
        self.last_word_path: Path | None = None
        self.last_validation_path: Path | None = None
        self.last_intake_preview_dir: Path | None = None
        self.analyzed_objects: list[dict[str, Any]] = []
        self.object_choices_by_label: dict[str, dict[str, Any]] = {}

        self.pdf_path = StringVar(value=_default_file("Діпр.наб. 17-К.pdf"))
        self.apartment = StringVar(value="353")
        self.object_choice = StringVar(value="")
        self.profile = StringVar(value="apartment")
        self.excel_template = StringVar(value=_default_file("ДН 15-Ж Аналоги Червень.xls"))
        self.report_template = StringVar(value=_default_file("config/report_templates/valuation_report_real_template.docx"))
        self.output_dir = StringVar(value="")
        self.complex_name = StringVar(value="")
        self.batch_all_objects = BooleanVar(value=True)
        self.include_full_screenshots = BooleanVar(value=False)
        self.show_excel = BooleanVar(value=False)
        self.status = StringVar(value="Готово")
        self.validation_status = StringVar(value="Перевірка звіту: не запускалась")

        self._configure_styles()
        self._build_ui()
        self.after(150, self._drain_messages)

    def _configure_styles(self) -> None:
        for font_name, size in [
            ("TkDefaultFont", 10),
            ("TkTextFont", 10),
            ("TkMenuFont", 10),
            ("TkHeadingFont", 11),
        ]:
            try:
                tkfont.nametofont(font_name).configure(family="Segoe UI", size=size)
            except Exception:
                pass

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=("Segoe UI", 10), background=PALETTE["app_bg"], foreground=PALETTE["text"])
        style.configure("Root.TFrame", background=PALETTE["app_bg"])
        style.configure("Header.TFrame", background=PALETTE["app_bg"])
        style.configure("Card.TFrame", background=PALETTE["card_bg"])
        style.configure("Card.TLabelframe", background=PALETTE["card_bg"], bordercolor=PALETTE["border"], relief="solid")
        style.configure(
            "Card.TLabelframe.Label",
            background=PALETTE["card_bg"],
            foreground=PALETTE["primary_dark"],
            font=("Segoe UI Semibold", 12),
        )
        style.configure("Title.TLabel", background=PALETTE["app_bg"], foreground=PALETTE["text"], font=("Segoe UI Semibold", 22))
        style.configure("Subtitle.TLabel", background=PALETTE["app_bg"], foreground=PALETTE["muted"], font=("Segoe UI", 11))
        style.configure("Field.TLabel", background=PALETTE["card_bg"], foreground=PALETTE["text"], font=("Segoe UI Semibold", 10))
        style.configure("Muted.TLabel", background=PALETTE["app_bg"], foreground=PALETTE["muted"], font=("Segoe UI", 9))
        style.configure("Status.TLabel", background=PALETTE["app_bg"], foreground=PALETTE["text"], font=("Segoe UI Semibold", 10))
        style.configure("TEntry", fieldbackground="#ffffff", bordercolor=PALETTE["border"], lightcolor=PALETTE["border"], padding=6)
        style.configure("TCombobox", fieldbackground="#ffffff", bordercolor=PALETTE["border"], lightcolor=PALETTE["border"], padding=6)
        style.configure("TCheckbutton", background=PALETTE["card_bg"], foreground=PALETTE["text"], font=("Segoe UI", 10))
        style.configure("TButton", padding=(14, 8), font=("Segoe UI Semibold", 10), borderwidth=1)
        style.configure("Primary.TButton", background=PALETTE["primary"], foreground="#ffffff", bordercolor=PALETTE["primary_dark"])
        style.map("Primary.TButton", background=[("active", PALETTE["primary_dark"]), ("disabled", "#b8c7dc")])
        style.configure("Secondary.TButton", background=PALETTE["secondary"], foreground=PALETTE["text"], bordercolor=PALETTE["border"])
        style.configure("Success.TButton", background="#e6f4ea", foreground=PALETTE["success"], bordercolor="#b7dfc2")
        style.configure("Tool.TButton", background="#f8fafc", foreground=PALETTE["primary_dark"], bordercolor=PALETTE["border"], padding=(12, 7))

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=20, style="Root.TFrame")
        root.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(3, weight=1)

        header = ttk.Frame(root, style="Header.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Realtify", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Формування звіту оцінки: PDF -> аналоги -> Excel -> Word -> validation",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(3, 0))

        form = ttk.LabelFrame(root, text="1. Вхідні файли та параметри", padding=16, style="Card.TLabelframe")
        form.grid(row=1, column=0, sticky="ew")
        form.columnconfigure(1, weight=1)

        row = 0
        self._path_row(
            form,
            row,
            "PDF витяг/техпаспорт",
            self.pdf_path,
            [("PDF", "*.pdf"), ("Всі файли", "*.*")],
            command=self._pick_pdf_file,
        )
        row += 1
        self._entry_row(form, row, "Номер квартири/об'єкта", self.apartment)
        row += 1
        self._object_choice_row(form, row)
        row += 1
        self._combo_row(form, row, "Тип об'єкта", self.profile, ["apartment", "parking", "commercial", "office", "retail", "warehouse", "house", "land"])
        row += 1
        self._path_row(form, row, "Excel шаблон", self.excel_template, [("Excel", "*.xls *.xlsx *.xlsm"), ("Всі файли", "*.*")])
        row += 1
        self._path_row(form, row, "Word шаблон", self.report_template, [("Word", "*.docx"), ("Всі файли", "*.*")])
        row += 1
        self._dir_row(form, row, "Папка результату", self.output_dir)
        row += 1
        self._entry_row(form, row, "Назва ЖК/комплексу", self.complex_name)

        options = ttk.LabelFrame(root, text="2. Керування та результат", padding=16, style="Card.TLabelframe")
        options.grid(row=2, column=0, sticky="ew", pady=(14, 14))
        options.columnconfigure(3, weight=1)
        ttk.Checkbutton(options, text="Сформувати звіти для всіх об'єктів PDF", variable=self.batch_all_objects).grid(row=0, column=0, sticky="w", padx=(0, 20))
        ttk.Checkbutton(options, text="Додати full-page screenshots у Word", variable=self.include_full_screenshots).grid(row=1, column=0, sticky="w", padx=(0, 20), pady=(8, 0))
        ttk.Checkbutton(options, text="Показувати Excel під час заповнення", variable=self.show_excel).grid(row=1, column=1, sticky="w", padx=(0, 20), pady=(8, 0))

        actions = ttk.Frame(options, style="Card.TFrame")
        actions.grid(row=0, column=4, rowspan=2, sticky="e")
        self.env_button = ttk.Button(actions, text="Перевірити середовище", command=self._check_environment, style="Secondary.TButton")
        self.env_button.grid(row=0, column=0, padx=(0, 8))
        self.run_button = ttk.Button(actions, text="Сформувати звіти", command=self._start_workflow, style="Primary.TButton")
        self.run_button.grid(row=0, column=1, padx=(0, 8))
        self.open_button = ttk.Button(actions, text="Відкрити папку", command=self._open_output_dir, state="disabled", style="Tool.TButton")
        self.open_button.grid(row=0, column=2, padx=(0, 8))
        self.open_word_button = ttk.Button(actions, text="Word", command=self._open_word_report, state="disabled", style="Tool.TButton")
        self.open_word_button.grid(row=0, column=3, padx=(0, 8))
        self.open_excel_button = ttk.Button(actions, text="Excel", command=self._open_excel_report, state="disabled", style="Tool.TButton")
        self.open_excel_button.grid(row=0, column=4, padx=(0, 8))
        self.open_validation_button = ttk.Button(actions, text="Перевірка", command=self._open_validation_report, state="disabled", style="Success.TButton")
        self.open_validation_button.grid(row=0, column=5)

        log_frame = ttk.LabelFrame(root, text="3. Журнал виконання", padding=12, style="Card.TLabelframe")
        log_frame.grid(row=3, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = Text(
            log_frame,
            height=18,
            wrap="word",
            state="disabled",
            bg=PALETTE["log_bg"],
            fg=PALETTE["log_fg"],
            insertbackground=PALETTE["log_fg"],
            relief="flat",
            padx=14,
            pady=12,
            font=("Consolas", 10),
        )
        self.log.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=scrollbar.set)

        status_bar = ttk.Frame(root, padding=(0, 10, 0, 0), style="Root.TFrame")
        status_bar.grid(row=4, column=0, sticky="ew")
        status_bar.columnconfigure(0, weight=1)
        ttk.Label(status_bar, textvariable=self.status, style="Status.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(status_bar, textvariable=self.validation_status, style="Status.TLabel").grid(row=0, column=1, sticky="w", padx=(20, 0))
        ttk.Label(status_bar, text=f"Проект: {PROJECT_ROOT}", style="Muted.TLabel").grid(row=0, column=2, sticky="e")

    def _path_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: StringVar,
        filetypes: list[tuple[str, str]],
        *,
        command=None,
    ) -> None:
        ttk.Label(parent, text=label, style="Field.TLabel").grid(row=row, column=0, sticky="w", pady=7, padx=(0, 12))
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=7)
        button_command = command or (lambda: self._pick_file(variable, filetypes))
        ttk.Button(parent, text="Обрати", command=button_command, style="Secondary.TButton").grid(row=row, column=2, sticky="e", pady=7, padx=(10, 0))

    def _dir_row(self, parent: ttk.Frame, row: int, label: str, variable: StringVar) -> None:
        ttk.Label(parent, text=label, style="Field.TLabel").grid(row=row, column=0, sticky="w", pady=7, padx=(0, 12))
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=7)
        ttk.Button(parent, text="Обрати", command=lambda: self._pick_dir(variable), style="Secondary.TButton").grid(row=row, column=2, sticky="e", pady=7, padx=(10, 0))

    def _entry_row(self, parent: ttk.Frame, row: int, label: str, variable: StringVar) -> None:
        ttk.Label(parent, text=label, style="Field.TLabel").grid(row=row, column=0, sticky="w", pady=7, padx=(0, 12))
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, columnspan=2, sticky="ew", pady=7)

    def _combo_row(self, parent: ttk.Frame, row: int, label: str, variable: StringVar, values: list[str]) -> None:
        ttk.Label(parent, text=label, style="Field.TLabel").grid(row=row, column=0, sticky="w", pady=7, padx=(0, 12))
        ttk.Combobox(parent, textvariable=variable, values=values, state="readonly").grid(row=row, column=1, columnspan=2, sticky="ew", pady=7)

    def _object_choice_row(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="Знайдені об'єкти PDF", style="Field.TLabel").grid(row=row, column=0, sticky="w", pady=7, padx=(0, 12))
        self.object_choice_box = ttk.Combobox(parent, textvariable=self.object_choice, values=[], state="disabled")
        self.object_choice_box.grid(row=row, column=1, columnspan=2, sticky="ew", pady=7)
        self.object_choice_box.bind("<<ComboboxSelected>>", lambda _event: self._apply_selected_object_choice())

    def _pick_file(self, variable: StringVar, filetypes: list[tuple[str, str]]) -> bool:
        initial = _initial_dir(variable.get())
        path = filedialog.askopenfilename(initialdir=initial, filetypes=filetypes)
        if path:
            variable.set(path)
            return True
        return False

    def _pick_pdf_file(self) -> None:
        changed = self._pick_file(self.pdf_path, [("PDF", "*.pdf"), ("Всі файли", "*.*")])
        if changed and self.pdf_path.get().strip():
            self._start_pdf_analysis()

    def _pick_dir(self, variable: StringVar) -> None:
        initial = _initial_dir(variable.get()) if variable.get().strip() else str(PROJECT_ROOT / "outputs")
        path = filedialog.askdirectory(initialdir=initial)
        if path:
            variable.set(path)

    def _check_environment(self) -> None:
        status = collect_status()
        missing_modules = [name for name, ok in status["modules"].items() if not ok]
        problems: list[str] = []
        if missing_modules:
            problems.append("Відсутні Python-модулі: " + ", ".join(missing_modules))
        if not status["tesseract"]:
            problems.append("Tesseract OCR не знайдено")
        if not status["poppler_bin"]:
            problems.append("Poppler не знайдено")
        langs = set(status["ocr_languages"])
        if not {"eng", "ukr", "rus", "osd"}.issubset(langs):
            problems.append("Не вистачає OCR-мов: eng, ukr, rus, osd")

        self._append_log("Перевірка середовища:")
        self._append_log(f"Python: {status['python']}")
        self._append_log(f"Tesseract: {status['tesseract'] or 'MISSING'}")
        self._append_log(f"Poppler: {status['poppler_bin'] or 'MISSING'}")
        self._append_log(f"OCR languages: {', '.join(status['ocr_languages']) or 'none'}")
        self._append_log(f"Playwright browsers: {status.get('playwright_browsers_dir') or 'default user cache'}")
        self._append_log(f"Bundled browser dirs: {', '.join(status.get('playwright_browsers') or []) or 'none'}")
        if problems:
            self.status.set("Середовище потребує налаштування")
            messagebox.showwarning("Перевірка середовища", "\n".join(problems))
        else:
            self.status.set("Середовище готове")
            messagebox.showinfo("Перевірка середовища", "Середовище готове до формування звіту.")

    def _start_workflow(self) -> None:
        if self.analysis_worker and self.analysis_worker.is_alive():
            messagebox.showinfo("Зачекайте", "Йде аналіз PDF. Дочекайтесь завершення або виберіть об'єкт вручну після аналізу.")
            return
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Зачекайте", "Формування звіту вже виконується.")
            return

        try:
            params = self._build_params()
        except ValueError as exc:
            messagebox.showerror("Некоректні дані", str(exc))
            return

        self.last_output_dir = params.output_dir
        self.last_excel_path = None
        self.last_word_path = None
        self.last_validation_path = None
        self._set_result_buttons_state("disabled")
        self.run_button.configure(state="disabled")
        self.env_button.configure(state="disabled")
        self.status.set("Формування звіту...")
        self.validation_status.set("Перевірка звіту: очікує завершення")
        self._clear_log()
        self._append_log("Запуск workflow:")
        self._append_log(f"PDF: {params.pdf_path}")
        self._append_log(f"Excel: {params.template_path}")
        self._append_log(f"Word template: {params.report_template_path}")
        self._append_log(f"Output: {params.output_dir}")
        if params.batch_all_objects:
            self._append_log("Mode: batch - звіт по кожному об'єкту з PDF")
        elif params.apartment:
            self._append_log(f"Object/apartment: {params.apartment}")
        self._append_log("Збір аналогів, OCR, скриншоти та Excel/Word можуть зайняти кілька хвилин.")
        self._append_log("")

        self.worker = threading.Thread(target=self._run_workflow, args=(params,), daemon=True)
        self.worker.start()

    def _build_params(self) -> WorkflowParams:
        pdf = _required_path(self.pdf_path.get(), "PDF витяг/техпаспорт")
        excel = _required_path(self.excel_template.get(), "Excel шаблон")
        report_template = _required_path(self.report_template.get(), "Word шаблон")
        profile = self.profile.get().strip() or "apartment"
        output = self.output_dir.get().strip()
        if output:
            output_dir = Path(output)
        else:
            output_dir = PROJECT_ROOT / "outputs" / f"app_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            self.output_dir.set(str(output_dir))

        apartment = self.apartment.get().strip()
        batch_all_objects = self.batch_all_objects.get()
        if not batch_all_objects and len(self.analyzed_objects) > 1 and not apartment:
            raise ValueError("У PDF знайдено кілька об'єктів. Оберіть потрібний об'єкт у полі 'Знайдені об'єкти PDF'.")
        complex_name = self.complex_name.get().strip()
        return WorkflowParams(
            pdf_path=pdf,
            template_path=excel,
            report_template_path=report_template,
            output_dir=output_dir,
            apartment=apartment or None,
            profile=profile,
            complex_name=complex_name or None,
            batch_all_objects=batch_all_objects,
            include_full_screenshots=self.include_full_screenshots.get(),
            visible=self.show_excel.get(),
        )

    def _start_pdf_analysis(self) -> None:
        if self.analysis_worker and self.analysis_worker.is_alive():
            self._append_log("Аналіз PDF вже виконується.")
            return
        if self.worker and self.worker.is_alive():
            self._append_log("Workflow вже виконується; автоматичний аналіз нового PDF пропущено.")
            return
        pdf_text = self.pdf_path.get().strip()
        if not pdf_text:
            return
        pdf = Path(pdf_text)
        if not pdf.exists():
            messagebox.showerror("Файл не знайдено", f"PDF не знайдено: {pdf}")
            return

        self.object_choice.set("")
        self.object_choices_by_label = {}
        self.analyzed_objects = []
        self.object_choice_box.configure(values=[], state="disabled")
        self.status.set("Аналіз PDF...")
        self.run_button.configure(state="disabled")
        self._append_log("")
        self._append_log(f"Старт аналізу PDF: {pdf}")
        self.analysis_worker = threading.Thread(target=self._run_pdf_analysis, args=(pdf,), daemon=True)
        self.analysis_worker.start()

    def _run_pdf_analysis(self, pdf: Path) -> None:
        preview_dir = PROJECT_ROOT / "outputs" / "_intake_preview" / datetime.now().strftime("%Y%m%d_%H%M%S")

        def progress(message: str) -> None:
            self.messages.put(("log", f"{datetime.now().strftime('%H:%M:%S')}  {message}"))

        try:
            template = Path(self.excel_template.get().strip()) if self.excel_template.get().strip() else None
            if template and not template.exists():
                template = None
            files = extract_intake_from_pdf(
                pdf_path=pdf,
                output_dir=preview_dir,
                target_apartment=self.apartment.get().strip() or None,
                template_path=template,
                profile=self.profile.get().strip() or "apartment",
                complex_name=self.complex_name.get().strip() or None,
                dpi=180,
                progress=progress,
            )
            self.messages.put(
                (
                    "analysis_done",
                    {
                        "preview_dir": preview_dir,
                        "objects": [record.model_dump(mode="json") for record in files.result.extracts],
                        "selected": files.result.selected_extract.model_dump(mode="json") if files.result.selected_extract else None,
                        "technical_count": len(files.result.technical_passports),
                        "warnings": files.result.warnings,
                    },
                )
            )
        except Exception as exc:
            self.messages.put(("analysis_failed", f"{exc}\n\n{traceback.format_exc()}"))

    def _apply_pdf_analysis(self, payload: dict[str, Any]) -> None:
        records = list(payload.get("objects") or [])
        self.last_intake_preview_dir = Path(payload["preview_dir"])
        self.analyzed_objects = records
        self.object_choices_by_label = {}
        labels: list[str] = []
        for record in records:
            label = _format_object_choice_label(record)
            labels.append(label)
            self.object_choices_by_label[label] = record

        if labels:
            self.object_choice_box.configure(values=labels, state="readonly")
        else:
            self.object_choice_box.configure(values=[], state="disabled")

        selected = payload.get("selected")
        selected_label = ""
        current_apartment = self.apartment.get().strip()
        for label, record in self.object_choices_by_label.items():
            if selected and record.get("page") == selected.get("page"):
                selected_label = label
                break
            if current_apartment and record.get("apartment_number") == current_apartment:
                selected_label = label
        if not selected_label and labels:
            selected_label = labels[0]
            if len(labels) > 1:
                self._append_log("У PDF знайдено кілька об'єктів. Автоматично обрано перший; перевірте список перед запуском звіту.")

        if selected_label:
            self.object_choice.set(selected_label)
            self._apply_selected_object_choice()

        self._append_log(f"Аналіз PDF завершено. Витягів знайдено: {len(records)}; техпаспортів: {payload.get('technical_count', 0)}.")
        self._append_log(f"Файли попереднього аналізу: {self.last_intake_preview_dir}")
        for warning in payload.get("warnings") or []:
            self._append_log(f"Попередження intake: {warning}")
        self.status.set("PDF проаналізовано")
        if not self.worker or not self.worker.is_alive():
            self.run_button.configure(state="normal")

    def _apply_selected_object_choice(self) -> None:
        record = self.object_choices_by_label.get(self.object_choice.get())
        if not record:
            return
        apartment = record.get("apartment_number")
        if apartment:
            self.apartment.set(str(apartment))
        area = record.get("total_area_m2")
        address = record.get("address_full") or ""
        self._append_log(
            "Обрано об'єкт з PDF: "
            f"стор. {record.get('page')}, "
            f"кв./об'єкт {apartment or 'не знайдено'}, "
            f"площа {area if area is not None else 'не знайдено'}, "
            f"{address}"
        )

    def _run_workflow(self, params: "WorkflowParams") -> None:
        pythoncom = None
        params.output_dir.mkdir(parents=True, exist_ok=True)
        runtime_log = params.output_dir / "runtime.log"

        def progress(message: str) -> None:
            line = f"{datetime.now().strftime('%H:%M:%S')}  {message}"
            try:
                with runtime_log.open("a", encoding="utf-8") as f:
                    f.write(line + "\n")
            except Exception:
                pass
            self.messages.put(("log", line))

        try:
            progress("Worker started.")
            if os.name == "nt":
                try:
                    import pythoncom as pythoncom_module  # type: ignore[import-not-found]

                    pythoncom = pythoncom_module
                    pythoncom.CoInitialize()
                    progress("Windows COM initialized.")
                except Exception:
                    pythoncom = None
                    progress("Windows COM initialization skipped or failed; continuing.")
            if params.batch_all_objects:
                result = run_batch_workflow(
                    pdf_path=params.pdf_path,
                    links_path=None,
                    template_path=params.template_path,
                    output_dir=params.output_dir,
                    profile=params.profile,
                    complex_name=params.complex_name,
                    report_template_path=params.report_template_path,
                    include_full_screenshots=params.include_full_screenshots,
                    visible=params.visible,
                    progress=progress,
                )
                passed = sum(1 for item in result.objects if item.ok)
                self.messages.put(
                    (
                        "done_batch",
                        {
                            "output_dir": result.output_dir,
                            "objects_total": len(result.objects),
                            "objects_ok": passed,
                            "objects_failed": len(result.objects) - passed,
                            "validation_ok": result.ok,
                            "batch_report": result.report_path,
                        },
                    )
                )
            else:
                result = run_full_workflow(
                    pdf_path=params.pdf_path,
                    links_path=None,
                    template_path=params.template_path,
                    output_dir=params.output_dir,
                    apartment=params.apartment,
                    profile=params.profile,
                    complex_name=params.complex_name,
                    report_template_path=params.report_template_path,
                    include_full_screenshots=params.include_full_screenshots,
                    visible=params.visible,
                    progress=progress,
                )
                raw_count = len(result.excel_workflow.raw_collection.candidates) if result.excel_workflow.raw_collection else len(result.excel_workflow.collection.candidates)
                self.messages.put(
                    (
                        "done",
                        {
                            "output_dir": result.output_dir,
                            "excel": result.excel_workflow.excel.output_path if result.excel_workflow.excel else None,
                            "word": result.word_report_path,
                            "collected": raw_count,
                            "selected": len(result.excel_workflow.collection.candidates),
                            "errors": len(result.excel_workflow.collection.errors),
                            "validation_ok": result.validation.ok if result.validation else None,
                            "validation_report": result.validation.validation_md if result.validation else None,
                        },
                    )
                )
        except Exception as exc:
            progress(f"Workflow failed: {exc}")
            self.messages.put(("failed", f"{exc}\n\n{traceback.format_exc()}"))
        finally:
            if pythoncom:
                try:
                    pythoncom.CoUninitialize()
                    self.messages.put(("log", f"{datetime.now().strftime('%H:%M:%S')}  Windows COM finalized."))
                except Exception:
                    pass

    def _drain_messages(self) -> None:
        try:
            while True:
                kind, payload = self.messages.get_nowait()
                if kind == "log":
                    self._append_log(payload)
                elif kind == "done":
                    output_dir = Path(payload["output_dir"])
                    self.last_output_dir = output_dir
                    self.last_excel_path = Path(payload["excel"]) if payload.get("excel") else None
                    self.last_word_path = Path(payload["word"]) if payload.get("word") else None
                    self.last_validation_path = Path(payload["validation_report"]) if payload.get("validation_report") else None
                    self.status.set("Звіт сформовано")
                    if payload.get("validation_ok") is True:
                        self.validation_status.set("Перевірка звіту: PASS")
                    elif payload.get("validation_ok") is False:
                        self.validation_status.set("Перевірка звіту: FAIL")
                    else:
                        self.validation_status.set("Перевірка звіту: не запускалась")
                    self.run_button.configure(state="normal")
                    self.env_button.configure(state="normal")
                    self._refresh_result_buttons()
                    self._append_log("")
                    self._append_log(f"Готово. Папка результату: {output_dir}")
                    self._append_log(f"Зібрано оголошень: {payload['collected']}")
                    self._append_log(f"Відібрано аналогів: {payload['selected']}")
                    self._append_log(f"Помилок збору: {payload['errors']}")
                    if payload.get("excel"):
                        self._append_log(f"Excel: {payload['excel']}")
                    if payload.get("word"):
                        self._append_log(f"Word: {payload['word']}")
                    if payload.get("validation_report"):
                        self._append_log(f"Validation: {'PASS' if payload.get('validation_ok') else 'FAIL'}")
                        self._append_log(f"Validation report: {payload['validation_report']}")
                    messagebox.showinfo("Готово", f"Звіт сформовано.\n\nПапка результату:\n{output_dir}")
                elif kind == "done_batch":
                    output_dir = Path(payload["output_dir"])
                    self.last_output_dir = output_dir
                    self.last_excel_path = None
                    self.last_word_path = None
                    self.last_validation_path = Path(payload["batch_report"]) if payload.get("batch_report") else None
                    self.status.set("Пакет звітів сформовано")
                    self.validation_status.set(
                        "Перевірка пакета: PASS"
                        if payload.get("validation_ok")
                        else f"Перевірка пакета: FAIL ({payload.get('objects_failed', 0)} пом.)"
                    )
                    self.run_button.configure(state="normal")
                    self.env_button.configure(state="normal")
                    self._refresh_result_buttons()
                    self._append_log("")
                    self._append_log(f"Готово. Папка пакета: {output_dir}")
                    self._append_log(f"Об'єктів у PDF: {payload['objects_total']}")
                    self._append_log(f"Звітів PASS: {payload['objects_ok']}")
                    self._append_log(f"Звітів FAIL: {payload['objects_failed']}")
                    if payload.get("batch_report"):
                        self._append_log(f"Batch report: {payload['batch_report']}")
                    messagebox.showinfo(
                        "Готово",
                        "Пакет звітів сформовано.\n\n"
                        f"PASS: {payload['objects_ok']}/{payload['objects_total']}\n\n"
                        f"Папка результату:\n{output_dir}",
                    )
                elif kind == "failed":
                    self.status.set("Помилка формування звіту")
                    self.validation_status.set("Перевірка звіту: не пройдена")
                    self.run_button.configure(state="normal")
                    self.env_button.configure(state="normal")
                    if self.last_output_dir:
                        self._refresh_result_buttons()
                    self._append_log("")
                    self._append_log(payload)
                    messagebox.showerror("Помилка", payload)
                elif kind == "analysis_done":
                    self._apply_pdf_analysis(payload)
                elif kind == "analysis_failed":
                    self.status.set("Помилка аналізу PDF")
                    if not self.worker or not self.worker.is_alive():
                        self.run_button.configure(state="normal")
                    self._append_log("")
                    self._append_log(payload)
                    messagebox.showerror("Помилка аналізу PDF", payload)
        except queue.Empty:
            pass
        self.after(150, self._drain_messages)

    def _open_output_dir(self) -> None:
        path = self.last_output_dir or Path(self.output_dir.get())
        if not path:
            return
        path.mkdir(parents=True, exist_ok=True)
        self._open_path(path)

    def _open_word_report(self) -> None:
        self._open_path(self.last_word_path)

    def _open_excel_report(self) -> None:
        self._open_path(self.last_excel_path)

    def _open_validation_report(self) -> None:
        self._open_path(self.last_validation_path)

    def _open_path(self, path: Path | None) -> None:
        if not path or not path.exists():
            messagebox.showwarning("Файл не знайдено", "Результат ще не створено або файл відсутній.")
            return
        os.startfile(path)  # type: ignore[attr-defined]

    def _set_result_buttons_state(self, state: str) -> None:
        self.open_button.configure(state=state)
        self.open_word_button.configure(state=state)
        self.open_excel_button.configure(state=state)
        self.open_validation_button.configure(state=state)

    def _refresh_result_buttons(self) -> None:
        self.open_button.configure(state="normal" if self.last_output_dir else "disabled")
        self.open_word_button.configure(state="normal" if self.last_word_path and self.last_word_path.exists() else "disabled")
        self.open_excel_button.configure(state="normal" if self.last_excel_path and self.last_excel_path.exists() else "disabled")
        self.open_validation_button.configure(state="normal" if self.last_validation_path and self.last_validation_path.exists() else "disabled")

    def _append_log(self, line: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", line + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")


def _default_file(relative: str) -> str:
    for root in (PROJECT_ROOT, RESOURCE_ROOT):
        path = root / relative
        if path.exists():
            return str(path)
    return ""


def _initial_dir(value: str) -> str:
    if not value:
        return str(PROJECT_ROOT)
    path = Path(value)
    if path.is_file():
        return str(path.parent)
    if path.is_dir():
        return str(path)
    return str(PROJECT_ROOT)


def _format_object_choice_label(record: dict[str, Any]) -> str:
    page = record.get("page") or "?"
    apartment = record.get("apartment_number") or "без номера"
    area = record.get("total_area_m2")
    area_text = f"{area} м²" if area is not None else "площа ?"
    address = str(record.get("address_full") or "").strip()
    if len(address) > 90:
        address = address[:87].rstrip() + "..."
    return f"стор. {page} | кв./об'єкт {apartment} | {area_text} | {address}"


def _required_path(value: str, label: str) -> Path:
    if not value.strip():
        raise ValueError(f"Поле '{label}' обов'язкове.")
    path = Path(value.strip())
    if not path.exists():
        raise ValueError(f"Файл не знайдено: {path}")
    return path


@dataclass(frozen=True)
class WorkflowParams:
    pdf_path: Path
    template_path: Path
    report_template_path: Path
    output_dir: Path
    apartment: str | None
    profile: str
    complex_name: str | None
    batch_all_objects: bool
    include_full_screenshots: bool
    visible: bool


if __name__ == "__main__":
    raise SystemExit(main())
