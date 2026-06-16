from __future__ import annotations

import argparse
import base64
import binascii
import hashlib
import hmac
import html
import json
import os
import re
import secrets
import shutil
import threading
import time
import traceback
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import quote
from uuid import uuid4

import requests
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles

from realtify.batch_workflow import BatchWorkflowResult, run_batch_workflow
from realtify.excel_tools import excel_com_available, libreoffice_available
from realtify.paths import PROJECT_ROOT


WEB_RUNS_ROOT = PROJECT_ROOT / "web_runs"
STATIC_DIR = Path(__file__).resolve().parent / "static"
DEFAULT_WORD_TEMPLATE = PROJECT_ROOT / "config" / "report_templates" / "valuation_report_real_template.docx"
AUTH_COOKIE_NAME = "autovalue_session"
AUTH_SESSION_SECONDS = 12 * 60 * 60

JOBS_LOCK = threading.RLock()
JOBS: dict[str, "WebJob"] = {}


@dataclass
class WebArtifact:
    name: str
    path: Path
    kind: str


@dataclass
class WebJob:
    id: str
    status: str
    created_at: str
    input_dir: Path
    output_dir: Path
    runtime_log: Path
    started_at: str | None = None
    finished_at: str | None = None
    error: str | None = None
    cancel_requested: bool = False
    events: list[str] = field(default_factory=list)
    artifacts: list[WebArtifact] = field(default_factory=list)


class JobCancelled(RuntimeError):
    pass


app = FastAPI(title="Realtify Web", version="0.1.0")
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.middleware("http")
async def auth_middleware(request: Request, call_next: Any) -> Response:
    credentials = _auth_credentials()
    if credentials is None:
        return await call_next(request)

    if request.url.path in {"/login", "/logout"}:
        return await call_next(request)

    if _session_token_valid(request.cookies.get(AUTH_COOKIE_NAME), username=credentials[0], password=credentials[1]):
        return await call_next(request)

    if request.url.path.startswith("/api/"):
        return Response("Authentication required", status_code=401)

    return RedirectResponse(url=f"/login?next={quote(_next_url_from_request(request), safe='')}", status_code=303)


@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request) -> Response:
    credentials = _auth_credentials()
    next_url = _safe_next_url(request.query_params.get("next"))
    if credentials is None:
        return RedirectResponse(url=next_url, status_code=303)
    if _session_token_valid(request.cookies.get(AUTH_COOKIE_NAME), username=credentials[0], password=credentials[1]):
        return RedirectResponse(url=next_url, status_code=303)
    return HTMLResponse(_render_login_page(next_url=next_url))


@app.post("/login", response_class=HTMLResponse)
def login_submit(
    request: Request,
    username: str = Form(""),
    password: str = Form(""),
    next_url: str = Form("/"),
) -> Response:
    credentials = _auth_credentials()
    next_url = _safe_next_url(next_url)
    if credentials is None:
        return RedirectResponse(url=next_url, status_code=303)

    expected_username, expected_password = credentials
    if not (
        hmac.compare_digest(username, expected_username)
        and hmac.compare_digest(password, expected_password)
    ):
        return HTMLResponse(_render_login_page(next_url=next_url, error="Неверный логин или пароль."))

    response = RedirectResponse(url=next_url, status_code=303)
    response.set_cookie(
        AUTH_COOKIE_NAME,
        _make_session_token(username=expected_username, password=expected_password),
        max_age=AUTH_SESSION_SECONDS,
        httponly=True,
        secure=_secure_cookie(request),
        samesite="lax",
    )
    return response


@app.post("/logout")
def logout() -> Response:
    response = RedirectResponse(url="/login", status_code=303)
    response.delete_cookie(AUTH_COOKIE_NAME)
    return response


@app.get("/", response_class=HTMLResponse)
def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/api/health")
def health() -> dict[str, Any]:
    return {
        "ok": True,
        "project_root": str(PROJECT_ROOT),
        "web_runs_root": str(WEB_RUNS_ROOT),
        "excel_com_available": excel_com_available(),
        "libreoffice_backend_available": libreoffice_available(),
        "python_xls_backend_available": _python_xls_backend_available(),
        "calculation_available": excel_com_available() or libreoffice_available() or _python_xls_backend_available(),
        "platform": os.name,
        "default_word_template_exists": DEFAULT_WORD_TEMPLATE.exists(),
    }


@app.post("/api/jobs")
async def create_job(
    pdf_file: UploadFile = File(...),
    excel_template: UploadFile = File(...),
    word_template: UploadFile | None = File(None),
    links_file: UploadFile | None = File(None),
    profile: str = Form("apartment"),
    complex_name: str | None = Form(None),
    include_full_screenshots: bool = Form(False),
    first_page: int | None = Form(None),
    last_page: int | None = Form(None),
    required_count: int = Form(5),
    allow_less: bool = Form(False),
) -> dict[str, Any]:
    job_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid4().hex[:8]
    job_dir = WEB_RUNS_ROOT / job_id
    input_dir = job_dir / "input"
    output_dir = job_dir / "output"
    runtime_log = job_dir / "runtime.log"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_path = input_dir / _safe_upload_name(pdf_file.filename, "input.pdf")
    excel_path = input_dir / _safe_upload_name(excel_template.filename, "template.xls")
    word_path: Path | None = None
    links_path: Path | None = None

    await _save_upload(pdf_file, pdf_path)
    await _save_upload(excel_template, excel_path)
    if word_template and word_template.filename:
        word_path = input_dir / _safe_upload_name(word_template.filename, "report_template.docx")
        await _save_upload(word_template, word_path)
    if links_file and links_file.filename:
        links_path = input_dir / _safe_upload_name(links_file.filename, "links.txt")
        await _save_upload(links_file, links_path)

    if word_path is None and DEFAULT_WORD_TEMPLATE.exists():
        word_path = DEFAULT_WORD_TEMPLATE

    if word_path is None or not word_path.exists():
        raise HTTPException(status_code=400, detail="Word template is required and default template is missing.")

    job = WebJob(
        id=job_id,
        status="queued",
        created_at=_now_iso(),
        input_dir=input_dir,
        output_dir=output_dir,
        runtime_log=runtime_log,
    )
    with JOBS_LOCK:
        JOBS[job_id] = job
    _append_event(job_id, "Задача создана. Файлы сохранены.")

    thread = threading.Thread(
        target=_run_job,
        args=(
            job_id,
            pdf_path,
            excel_path,
            word_path,
            links_path,
            profile,
            complex_name,
            include_full_screenshots,
            first_page,
            last_page,
            max(1, required_count),
            allow_less,
        ),
        daemon=True,
        name=f"realtify-web-job-{job_id}",
    )
    thread.start()
    return _job_payload(job_id)


@app.post("/api/library/import")
async def import_library_job(library_file: UploadFile = File(...)) -> dict[str, Any]:
    job_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_lib_" + uuid4().hex[:6]
    job_dir = WEB_RUNS_ROOT / job_id
    input_dir = job_dir / "input"
    output_dir = job_dir / "output"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    file_path = input_dir / _safe_upload_name(library_file.filename, "library.csv")
    await _save_upload(library_file, file_path)

    job = WebJob(
        id=job_id,
        status="queued",
        created_at=_now_iso(),
        input_dir=input_dir,
        output_dir=output_dir,
        runtime_log=job_dir / "runtime.log",
    )
    with JOBS_LOCK:
        JOBS[job_id] = job
    _append_event(job_id, "Імпорт бібліотеки аналогів: файл збережено.")
    thread = threading.Thread(
        target=_run_import_job, args=(job_id, file_path), daemon=True, name=f"realtify-lib-{job_id}"
    )
    thread.start()
    return _job_payload(job_id)


def _run_import_job(job_id: str, file_path: Path) -> None:
    from realtify.analog_library import import_library, parse_library_file

    _set_job(job_id, status="running", started_at=_now_iso())
    try:
        entries = parse_library_file(file_path)
        _append_event(job_id, f"Імпорт бібліотеки: адрес у файлі — {len(entries)}.")
        if not entries:
            _append_event(job_id, "Помилка: у файлі немає валідних рядків address + url.")
            _set_job(
                job_id,
                status="failed",
                finished_at=_now_iso(),
                error="У файлі не знайдено рядків address + url.",
            )
            _notify_job_finished(job_id)
            return
        report = import_library(
            entries,
            output_dir=_get_job(job_id).output_dir,
            progress=lambda message: _append_event(job_id, message),
        )
        report_path = _get_job(job_id).output_dir / "library_import_report.json"
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        artifacts = [
            WebArtifact(name="library_import_report.json", path=report_path, kind="report")
        ]
        _set_job(job_id, status="passed", finished_at=_now_iso(), artifacts=artifacts)
        _append_event(
            job_id,
            f"Готово: збережено {report['saved_addresses']}/{report['addresses']} адрес, "
            f"{report['saved_analogs']} аналогів у бібліотеку.",
        )
        _notify_job_finished(job_id)
    except Exception as exc:  # noqa: BLE001 - web job must capture the full failure for the UI.
        _append_event(job_id, f"Помилка імпорту: {exc}")
        _append_event(job_id, traceback.format_exc())
        _set_job(job_id, status="failed", finished_at=_now_iso(), error=str(exc))
        _notify_job_finished(job_id)


@app.get("/api/jobs")
def list_jobs() -> dict[str, Any]:
    _hydrate_recent_jobs_from_disk()
    with JOBS_LOCK:
        ids = sorted(JOBS, reverse=True)
    return {"jobs": [_job_payload(job_id) for job_id in ids[:50]]}


@app.get("/api/jobs/{job_id}")
def get_job(job_id: str) -> dict[str, Any]:
    return _job_payload(job_id)


@app.post("/api/jobs/{job_id}/cancel")
def cancel_job(job_id: str) -> dict[str, Any]:
    job = _get_job(job_id)
    with JOBS_LOCK:
        if job.status in {"passed", "failed", "cancelled"}:
            return _job_payload(job_id)
        job.cancel_requested = True
        job.status = "cancelling"
        job.error = "Задача остановлена пользователем."
    _append_event(job_id, "Задача получила команду остановки. Workflow завершится на ближайшем безопасном шаге.")
    return _job_payload(job_id)


@app.get("/api/jobs/{job_id}/events")
def get_job_events(job_id: str) -> dict[str, Any]:
    job = _get_job(job_id)
    with JOBS_LOCK:
        events = list(job.events)
    return {"job_id": job_id, "events": events}


@app.get("/api/jobs/{job_id}/files/{artifact_name}")
def download_artifact(job_id: str, artifact_name: str) -> FileResponse:
    job = _get_job(job_id)
    with JOBS_LOCK:
        artifact = next((item for item in job.artifacts if item.name == artifact_name), None)
    if artifact is None or not artifact.path.exists():
        raise HTTPException(status_code=404, detail="Artifact not found.")
    return FileResponse(artifact.path, filename=artifact.name)


def _run_job(
    job_id: str,
    pdf_path: Path,
    excel_path: Path,
    word_path: Path,
    links_path: Path | None,
    profile: str,
    complex_name: str | None,
    include_full_screenshots: bool,
    first_page: int | None,
    last_page: int | None,
    required_count: int,
    allow_less: bool,
) -> None:
    if _job_cancel_requested(job_id):
        _mark_job_cancelled(job_id)
        return
    _set_job(job_id, status="running", started_at=_now_iso())
    _append_event(job_id, "Запуск batch workflow: один отчет на каждый объект из PDF.")
    _append_event(job_id, f"PDF: {pdf_path.name}")
    _append_event(job_id, f"Excel template: {excel_path.name}")
    _append_event(job_id, f"Word template: {word_path.name}")
    _append_event(job_id, f"Profile: {profile}; required analogs: {required_count}")
    _append_event(job_id, f"PDF render DPI: {_web_pdf_dpi()}")
    if first_page or last_page:
        _append_event(job_id, f"Ограничение страниц PDF: {first_page or 1}-{last_page or 'end'}")
    if not excel_com_available():
        if libreoffice_available():
            _append_event(job_id, "Excel COM недоступен; используется формуло-безопасный LibreOffice backend.")
        else:
            _append_event(job_id, "Excel COM и LibreOffice недоступны; используется аварийный Python XLS backend, который может заменить формулы значениями.")

    try:
        result = run_batch_workflow(
            pdf_path=pdf_path,
            links_path=links_path,
            template_path=excel_path,
            output_dir=_get_job(job_id).output_dir,
            profile=profile,
            complex_name=complex_name or None,
            required_count=required_count,
            allow_less=allow_less,
            allow_incomplete=False,
            first_page=first_page,
            last_page=last_page,
            dpi=_web_pdf_dpi(),
            visible=False,
            report_template_path=word_path,
            include_full_screenshots=include_full_screenshots,
            progress=lambda message: _append_progress_or_cancel(job_id, message),
        )
        _raise_if_cancelled(job_id)
        artifacts = _package_result(job_id, result)
        status = "passed" if result.ok else "failed"
        error = None if result.ok else "Один или несколько отчетов завершились с ошибкой. Проверьте batch_report.md."
        _set_job(job_id, status=status, finished_at=_now_iso(), error=error, artifacts=artifacts)
        _append_event(job_id, "Задача завершена: " + ("PASS" if result.ok else "FAIL"))
        _notify_job_finished(job_id)
    except JobCancelled:
        _mark_job_cancelled(job_id)
    except Exception as exc:  # noqa: BLE001 - web job must capture the full failure for the UI.
        trace = traceback.format_exc()
        _append_event(job_id, f"Ошибка: {exc}")
        _append_event(job_id, trace)
        _set_job(job_id, status="failed", finished_at=_now_iso(), error=str(exc))
        _notify_job_finished(job_id)


def _package_result(job_id: str, result: BatchWorkflowResult) -> list[WebArtifact]:
    job = _get_job(job_id)
    review_dir = job.output_dir / "client_review"
    review_dir.mkdir(parents=True, exist_ok=True)
    artifacts: list[WebArtifact] = []

    if result.report_path.exists():
        batch_report_copy = review_dir / "batch_report.md"
        shutil.copy2(result.report_path, batch_report_copy)
        artifacts.append(WebArtifact(name="batch_report.md", path=batch_report_copy, kind="report"))

    for index, item in enumerate(result.objects, start=1):
        label = _object_label(index, item.apartment, item.extract_page)
        if item.word_report_path and item.word_report_path.exists():
            target = review_dir / f"valuation_report_{label}.docx"
            shutil.copy2(item.word_report_path, target)
            artifacts.append(WebArtifact(name=target.name, path=target, kind="word"))
        if item.excel_workflow and item.excel_workflow.excel and item.excel_workflow.excel.output_path.exists():
            target = review_dir / f"valuation_calculation_{label}{item.excel_workflow.excel.output_path.suffix}"
            shutil.copy2(item.excel_workflow.excel.output_path, target)
            artifacts.append(WebArtifact(name=target.name, path=target, kind="excel"))
        if item.validation and item.validation.validation_md.exists():
            target = review_dir / f"validation_{label}.md"
            shutil.copy2(item.validation.validation_md, target)
            artifacts.append(WebArtifact(name=target.name, path=target, kind="validation"))

    zip_path = job.output_dir / "client_review_reports.zip"
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(review_dir.rglob("*")):
            if path.is_file():
                archive.write(path, arcname=path.relative_to(review_dir))
    artifacts.insert(0, WebArtifact(name=zip_path.name, path=zip_path, kind="zip"))
    _append_event(job_id, f"Архив для проверки: {zip_path}")
    return artifacts


async def _save_upload(upload: UploadFile, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open("wb") as handle:
        shutil.copyfileobj(upload.file, handle)


def _safe_upload_name(filename: str | None, fallback: str) -> str:
    name = Path(filename or fallback).name
    name = re.sub(r"[^A-Za-z0-9А-Яа-яЇїІіЄєҐґ._ -]+", "_", name, flags=re.UNICODE).strip()
    return name or fallback


def _object_label(index: int, apartment: str | None, page: int | None) -> str:
    raw = apartment or (f"page_{page}" if page else f"object_{index}")
    slug = re.sub(r"[^A-Za-z0-9_-]+", "-", raw).strip("-")
    return f"apt_{slug}" if apartment else slug


def _python_xls_backend_available() -> bool:
    try:
        import xlrd  # noqa: F401
        import xlutils.copy  # noqa: F401
        import xlwt  # noqa: F401
    except Exception:
        return False
    return True


def _web_pdf_dpi() -> int:
    raw_value = os.getenv("REALTIFY_WEB_PDF_DPI")
    if raw_value:
        try:
            value = int(raw_value)
            if 50 <= value <= 300:
                return value
        except ValueError:
            pass
    return 110


def _notify_job_finished(job_id: str) -> None:
    if not _telegram_enabled():
        return
    try:
        _send_telegram_message(_telegram_message(job_id))
        _append_event(job_id, "Telegram: уведомление о завершении отправлено.")
    except Exception as exc:  # noqa: BLE001 - notification failure must not break report generation.
        _append_event(job_id, f"Telegram: не удалось отправить уведомление: {exc}")


def _telegram_enabled() -> bool:
    return bool(os.getenv("REALTIFY_TELEGRAM_BOT_TOKEN") and os.getenv("REALTIFY_TELEGRAM_CHAT_ID"))


def _telegram_message(job_id: str) -> str:
    job = _get_job(job_id)
    with JOBS_LOCK:
        status = job.status.upper()
        started_at = job.started_at or "-"
        finished_at = job.finished_at or "-"
        error = job.error
        artifacts = list(job.artifacts)

    lines = [
        "Autovalue: отчет завершен",
        f"Job: {job_id}",
        f"Status: {status}",
        f"Started: {started_at}",
        f"Finished: {finished_at}",
    ]
    zip_artifact = next((item for item in artifacts if item.name.endswith(".zip")), None)
    base_url = (os.getenv("REALTIFY_PUBLIC_BASE_URL") or "").rstrip("/")
    if zip_artifact and base_url:
        lines.append(f"Archive: {base_url}/api/jobs/{job_id}/files/{quote(zip_artifact.name)}")
    if error:
        lines.append(f"Error: {error}")
    return "\n".join(lines)


def _send_telegram_message(message: str) -> None:
    token = os.getenv("REALTIFY_TELEGRAM_BOT_TOKEN")
    chat_id = os.getenv("REALTIFY_TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        return
    response = requests.post(
        f"https://api.telegram.org/bot{token}/sendMessage",
        data={
            "chat_id": chat_id,
            "text": message,
            "disable_web_page_preview": "true",
        },
        timeout=12,
    )
    response.raise_for_status()


def _auth_credentials() -> tuple[str, str] | None:
    username = os.getenv("REALTIFY_WEB_USERNAME")
    password = os.getenv("REALTIFY_WEB_PASSWORD")
    if not username or not password:
        return None
    return username, password


def _make_session_token(*, username: str, password: str) -> str:
    issued_at = str(int(time.time()))
    nonce = secrets.token_urlsafe(12)
    body = f"{username}:{issued_at}:{nonce}"
    signature = hmac.new(_auth_secret(username=username, password=password), body.encode("utf-8"), hashlib.sha256).hexdigest()
    return base64.urlsafe_b64encode(f"{body}:{signature}".encode("utf-8")).decode("ascii").rstrip("=")


def _session_token_valid(token: str | None, *, username: str, password: str) -> bool:
    if not token:
        return False
    try:
        padding = "=" * (-len(token) % 4)
        payload = base64.urlsafe_b64decode((token + padding).encode("ascii")).decode("utf-8")
    except (binascii.Error, UnicodeDecodeError, ValueError):
        return False
    parts = payload.split(":")
    if len(parts) != 4:
        return False
    supplied_username, issued_at_text, nonce, supplied_signature = parts
    if not nonce or not hmac.compare_digest(supplied_username, username):
        return False
    try:
        issued_at = int(issued_at_text)
    except ValueError:
        return False
    now = int(time.time())
    if issued_at > now + 60 or now - issued_at > AUTH_SESSION_SECONDS:
        return False
    body = f"{supplied_username}:{issued_at_text}:{nonce}"
    expected_signature = hmac.new(_auth_secret(username=username, password=password), body.encode("utf-8"), hashlib.sha256).hexdigest()
    return hmac.compare_digest(supplied_signature, expected_signature)


def _auth_secret(*, username: str, password: str) -> bytes:
    return hashlib.sha256(f"{username}\0{password}".encode("utf-8")).digest()


def _secure_cookie(request: Request) -> bool:
    scheme = request.headers.get("x-forwarded-proto") or request.url.scheme
    return scheme == "https"


def _next_url_from_request(request: Request) -> str:
    next_url = request.url.path
    if request.url.query:
        next_url += f"?{request.url.query}"
    return _safe_next_url(next_url)


def _safe_next_url(value: str | None) -> str:
    if not value or not value.startswith("/") or value.startswith("//"):
        return "/"
    return value


def _render_login_page(*, next_url: str, error: str | None = None) -> str:
    safe_next = html.escape(_safe_next_url(next_url), quote=True)
    error_html = ""
    if error:
        error_html = f'<div class="alert">{html.escape(error)}</div>'
    return f"""<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Autovalue - вход</title>
  <style>
    :root {{
      color-scheme: light;
      --page: #edf3f8;
      --surface: #ffffff;
      --surface-soft: #f6f9fc;
      --line: #c6d4e1;
      --ink: #14212d;
      --muted: #607487;
      --brand: #145da0;
      --brand-strong: #0f4577;
      --accent: #0f766e;
      --danger: #b42318;
      --focus: rgba(20, 93, 160, 0.2);
    }}

    * {{ box-sizing: border-box; }}

    body {{
      margin: 0;
      min-width: 320px;
      min-height: 100vh;
      display: grid;
      place-items: center;
      background: var(--page);
      color: var(--ink);
      font-family: "Segoe UI", Arial, sans-serif;
    }}

    .login-shell {{
      width: min(960px, calc(100% - 32px));
      display: grid;
      grid-template-columns: minmax(280px, 0.85fr) minmax(320px, 1fr);
      min-height: 520px;
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
      background: var(--surface);
      box-shadow: 0 22px 60px rgba(16, 36, 56, 0.18);
    }}

    .brand-panel {{
      padding: 38px;
      background: #123f66;
      color: #fff;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
    }}

    .brand-name {{
      font-size: 28px;
      font-weight: 800;
      letter-spacing: 0;
    }}

    .brand-copy {{
      max-width: 340px;
    }}

    .brand-copy h1 {{
      margin: 0 0 14px;
      font-size: 32px;
      line-height: 1.12;
      letter-spacing: 0;
    }}

    .brand-copy p {{
      margin: 0;
      color: rgba(255, 255, 255, 0.78);
      line-height: 1.55;
    }}

    .status-line {{
      display: inline-flex;
      align-items: center;
      gap: 9px;
      color: rgba(255, 255, 255, 0.82);
      font-size: 14px;
    }}

    .status-dot {{
      width: 9px;
      height: 9px;
      border-radius: 999px;
      background: #2dd4bf;
    }}

    .form-panel {{
      display: grid;
      align-content: center;
      padding: 42px;
    }}

    .form-panel h2 {{
      margin: 0;
      font-size: 26px;
      line-height: 1.2;
      letter-spacing: 0;
    }}

    .form-panel .lead {{
      margin: 8px 0 26px;
      color: var(--muted);
      line-height: 1.5;
    }}

    form {{
      display: grid;
      gap: 16px;
    }}

    label {{
      display: grid;
      gap: 7px;
      font-weight: 700;
    }}

    input {{
      width: 100%;
      min-height: 44px;
      border: 1px solid #96aabd;
      border-radius: 6px;
      padding: 9px 11px;
      background: #fff;
      color: var(--ink);
      font: inherit;
    }}

    input:focus,
    button:focus {{
      outline: 3px solid var(--focus);
      outline-offset: 1px;
      border-color: var(--brand);
    }}

    button {{
      min-height: 45px;
      border: 1px solid var(--brand-strong);
      border-radius: 6px;
      padding: 10px 16px;
      background: var(--brand);
      color: #fff;
      cursor: pointer;
      font: inherit;
      font-weight: 750;
    }}

    button:hover {{
      background: var(--brand-strong);
    }}

    .alert {{
      margin: 0 0 18px;
      padding: 11px 12px;
      border: 1px solid rgba(180, 35, 24, 0.35);
      border-radius: 6px;
      background: #fff1f0;
      color: var(--danger);
      font-weight: 650;
    }}

    .meta {{
      margin-top: 18px;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.45;
    }}

    @media (max-width: 760px) {{
      body {{
        align-items: stretch;
      }}

      .login-shell {{
        width: 100%;
        min-height: 100vh;
        grid-template-columns: 1fr;
        border: 0;
        border-radius: 0;
      }}

      .brand-panel {{
        min-height: 230px;
        padding: 28px 24px;
      }}

      .brand-copy h1 {{
        font-size: 28px;
      }}

      .form-panel {{
        padding: 30px 24px;
      }}
    }}
  </style>
</head>
<body>
  <main class="login-shell">
    <section class="brand-panel" aria-label="Autovalue">
      <div class="brand-name">Autovalue</div>
      <div class="brand-copy">
        <h1>Формирование отчетов оценки</h1>
        <p>Закрытая рабочая панель для загрузки документов, подбора аналогов и подготовки отчетов.</p>
      </div>
      <div class="status-line"><span class="status-dot"></span><span>Доступ только для клиента</span></div>
    </section>
    <section class="form-panel">
      <div>
        <h2>Вход в систему</h2>
        <p class="lead">Введите логин и пароль, чтобы открыть рабочую панель.</p>
        {error_html}
        <form method="post" action="/login" autocomplete="on">
          <input type="hidden" name="next_url" value="{safe_next}">
          <label>
            Логин
            <input name="username" type="text" autocomplete="username" autofocus required>
          </label>
          <label>
            Пароль
            <input name="password" type="password" autocomplete="current-password" required>
          </label>
          <button type="submit">Войти</button>
        </form>
        <div class="meta">Сессия активна 12 часов. После завершения работы используйте кнопку выхода.</div>
      </div>
    </section>
  </main>
</body>
</html>"""


def _append_event(job_id: str, message: str) -> None:
    job = _get_job(job_id)
    line = f"[{datetime.now().strftime('%H:%M:%S')}] {message}"
    with JOBS_LOCK:
        job.events.append(line)
        job.runtime_log.parent.mkdir(parents=True, exist_ok=True)
        with job.runtime_log.open("a", encoding="utf-8") as handle:
            handle.write(line + "\n")


def _append_progress_or_cancel(job_id: str, message: str) -> None:
    _raise_if_cancelled(job_id)
    _append_event(job_id, message)
    _raise_if_cancelled(job_id)


def _job_cancel_requested(job_id: str) -> bool:
    job = _get_job(job_id)
    with JOBS_LOCK:
        return job.cancel_requested or job.status == "cancelling"


def _raise_if_cancelled(job_id: str) -> None:
    if _job_cancel_requested(job_id):
        raise JobCancelled("Job cancelled by user.")


def _mark_job_cancelled(job_id: str) -> None:
    job = _get_job(job_id)
    with JOBS_LOCK:
        already_cancelled = job.status == "cancelled"
        job.cancel_requested = True
        job.status = "cancelled"
        job.finished_at = job.finished_at or _now_iso()
        job.error = "Задача остановлена пользователем."
    if not already_cancelled:
        _append_event(job_id, "Задача отменена: CANCELLED")


def _set_job(job_id: str, **updates: Any) -> None:
    job = _get_job(job_id)
    with JOBS_LOCK:
        for key, value in updates.items():
            setattr(job, key, value)


def _get_job(job_id: str) -> WebJob:
    with JOBS_LOCK:
        job = JOBS.get(job_id)
    if job is None:
        job = _hydrate_job_from_disk(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job not found.")
    return job


def _hydrate_recent_jobs_from_disk(limit: int = 50) -> None:
    if not WEB_RUNS_ROOT.exists():
        return
    run_dirs = sorted((path for path in WEB_RUNS_ROOT.iterdir() if path.is_dir()), reverse=True)
    for path in run_dirs[:limit]:
        _hydrate_job_from_disk(path.name)


def _hydrate_job_from_disk(job_id: str) -> WebJob | None:
    with JOBS_LOCK:
        existing = JOBS.get(job_id)
    if existing is not None:
        return existing

    job_dir = WEB_RUNS_ROOT / job_id
    input_dir = job_dir / "input"
    output_dir = job_dir / "output"
    runtime_log = job_dir / "runtime.log"
    if not job_dir.is_dir() or not output_dir.exists():
        return None

    events = _read_runtime_events(runtime_log)
    status = _status_from_events(events)
    artifacts = _artifacts_from_disk(output_dir)
    finished_at = _mtime_iso(output_dir / "client_review_reports.zip") or _mtime_iso(runtime_log)
    error = None
    if status == "failed":
        error = "Один или несколько отчетов завершились с ошибкой. Проверьте batch_report.md."
    elif status == "cancelled":
        error = "Задача остановлена пользователем."

    job = WebJob(
        id=job_id,
        status=status,
        created_at=_created_at_from_job_id(job_id) or _mtime_iso(job_dir) or _now_iso(),
        input_dir=input_dir,
        output_dir=output_dir,
        runtime_log=runtime_log,
        started_at=_created_at_from_job_id(job_id),
        finished_at=finished_at if status in {"passed", "failed", "cancelled"} else None,
        error=error,
        events=events,
        artifacts=artifacts,
    )
    with JOBS_LOCK:
        JOBS[job_id] = job
    return job


def _read_runtime_events(runtime_log: Path) -> list[str]:
    if not runtime_log.exists():
        return []
    return runtime_log.read_text(encoding="utf-8", errors="replace").splitlines()


def _status_from_events(events: list[str]) -> str:
    for line in reversed(events):
        if "Задача отменена: CANCELLED" in line:
            return "cancelled"
        if "Задача получила команду остановки" in line:
            return "cancelled"
        if "Задача завершена: PASS" in line:
            return "passed"
        if "Задача завершена: FAIL" in line:
            return "failed"
    return "failed" if events else "queued"


def _artifacts_from_disk(output_dir: Path) -> list[WebArtifact]:
    artifacts: list[WebArtifact] = []
    zip_path = output_dir / "client_review_reports.zip"
    if zip_path.exists():
        artifacts.append(WebArtifact(name=zip_path.name, path=zip_path, kind="zip"))
    review_dir = output_dir / "client_review"
    if review_dir.exists():
        for path in sorted(review_dir.iterdir()):
            if path.is_file():
                artifacts.append(WebArtifact(name=path.name, path=path, kind=_artifact_kind(path)))
    return artifacts


def _artifact_kind(path: Path) -> str:
    if path.suffix.lower() == ".docx":
        return "word"
    if path.suffix.lower() in {".xls", ".xlsx", ".xlsm"}:
        return "excel"
    if path.name.startswith("validation_"):
        return "validation"
    if path.name == "batch_report.md":
        return "report"
    return "file"


def _created_at_from_job_id(job_id: str) -> str | None:
    match = re.match(r"(\d{8})_(\d{6})_", job_id)
    if not match:
        return None
    try:
        value = datetime.strptime("".join(match.groups()), "%Y%m%d%H%M%S")
    except ValueError:
        return None
    return value.isoformat(timespec="seconds")


def _mtime_iso(path: Path) -> str | None:
    if not path.exists():
        return None
    return datetime.fromtimestamp(path.stat().st_mtime).isoformat(timespec="seconds")


def _job_payload(job_id: str) -> dict[str, Any]:
    job = _get_job(job_id)
    with JOBS_LOCK:
        events = list(job.events[-250:])
        cancel_requested = job.cancel_requested
        artifacts = [
            {
                "name": item.name,
                "kind": item.kind,
                "url": f"/api/jobs/{job.id}/files/{item.name}",
                "size": item.path.stat().st_size if item.path.exists() else None,
            }
            for item in job.artifacts
        ]
    return {
        "id": job.id,
        "status": job.status,
        "cancel_requested": cancel_requested,
        "can_cancel": job.status in {"queued", "running", "cancelling"},
        "created_at": job.created_at,
        "started_at": job.started_at,
        "finished_at": job.finished_at,
        "error": job.error,
        "events": events,
        "artifacts": artifacts,
        "output_dir": str(job.output_dir),
    }


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run the Realtify web interface.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--reload", action="store_true")
    args = parser.parse_args(argv)

    import uvicorn

    uvicorn.run("realtify_web.app:app", host=args.host, port=args.port, reload=args.reload)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
