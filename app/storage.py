"""Filesystem-backed storage for clients and jobs.

Deliberately simple (JSON files on disk) for an internal single-instance
tool. Swapping this for a real database later is a contained change since
every other module only talks to the functions in this file.
"""
import json
import uuid
from datetime import datetime
from pathlib import Path

import openpyxl

DATA_DIR = Path("data")
PRACTICES_DIR = DATA_DIR / "practices"
CLIENTS_DIR = DATA_DIR / "clients"
JOBS_DIR = DATA_DIR / "jobs"

# Default customization config for a newly-uploaded template. A practice
# edits this to control which schedules get generated, where each one
# lands in the template, and the thresholds/cell conventions used - so
# adding a second (or fifth) template format later is configuration, not
# new code.
DEFAULT_TEMPLATE_CONFIG = {
    "schedules": {
        "tb_lead_schedule": {"enabled": True, "insert_after_sheet": None},
        "profit_and_loss": {"enabled": True, "insert_after_sheet": None},
        "balance_sheet": {"enabled": True, "insert_after_sheet": None},
        "corporation_tax": {"enabled": True, "insert_after_sheet": None},
        "fixed_asset_category": {"enabled": True, "insert_after_sheet": None},
        "fixed_asset_register": {"enabled": True, "insert_after_sheet": None},
        "control_account_rollforward": {"enabled": True, "insert_after_sheet": None},
        "nominal_matrix": {"enabled": True, "insert_after_sheet": None, "max_accounts": 6},
        "debtors_recon": {"enabled": True, "insert_after_sheet": None},
        "creditors_recon": {"enabled": True, "insert_after_sheet": None},
        "bank_recon": {"enabled": True, "insert_after_sheet": None},
        "vat_recon": {"enabled": True, "insert_after_sheet": None},
        "nominal_review": {"enabled": True, "insert_after_sheet": None},
        "points_forward": {"enabled": True, "insert_after_sheet": None},
    },
    "header_cells": {"client_name_cell": "A1", "period_cell": "A2", "schedule_title_cell": "A3"},
    "materiality": {"default_amount": 500, "variance_pct_threshold": 0.10},
    "numbering": {"style": "sequential", "start_at": 1},
}


def _new_id(prefix: str) -> str:
    return f"{prefix}_{uuid.uuid4().hex[:10]}"


def _read_json(path: Path) -> dict | None:
    if path.exists():
        return json.loads(path.read_text())
    return None


def _write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, default=str))


# ---------- practices ----------

def create_practice(name: str) -> dict:
    practice_id = _new_id("practice")
    practice = {"id": practice_id, "name": name, "created_at": datetime.utcnow().isoformat(), "default_template_id": None}
    _write_json(PRACTICES_DIR / practice_id / "info.json", practice)
    return practice


def get_practice(practice_id: str) -> dict | None:
    return _read_json(PRACTICES_DIR / practice_id / "info.json")


def save_practice(practice: dict) -> None:
    _write_json(PRACTICES_DIR / practice["id"] / "info.json", practice)


def list_practices() -> list[dict]:
    if not PRACTICES_DIR.exists():
        return []
    out = []
    for d in sorted(PRACTICES_DIR.iterdir()):
        info = _read_json(d / "info.json")
        if info:
            out.append(info)
    return out


# ---------- templates (scoped to a practice) ----------

def _normalise_template_file(file_path: Path) -> dict:
    """One-time cleanup at upload time (not repeated per job): loading and
    re-saving through openpyxl drops calcChain/printer-settings bloat and
    shrinks the file substantially - confirmed on a real 63-sheet template,
    24MB down to 11MB with cell data, formulas, and formatting all intact.
    It also drops embedded images and dropdown data-validation lists, which
    is the deliberate trade-off: doing this once at setup keeps every later
    job generation fast, at the cost of re-adding a logo/validation lists
    once per template rather than never. Falls back to the original file
    untouched if it can't be parsed."""
    original_size = file_path.stat().st_size
    try:
        wb = openpyxl.load_workbook(file_path)
        wb.save(file_path)
        normalised = True
    except Exception:
        normalised = False
    return {
        "normalised": normalised,
        "original_size_bytes": original_size,
        "stored_size_bytes": file_path.stat().st_size,
    }


def create_template(practice_id: str, name: str, file_bytes: bytes, filename: str) -> dict:
    template_id = _new_id("template")
    template_dir = PRACTICES_DIR / practice_id / "templates" / template_id
    template_dir.mkdir(parents=True, exist_ok=True)
    file_path = template_dir / filename
    file_path.write_bytes(file_bytes)
    normalisation = _normalise_template_file(file_path)

    template = {
        "id": template_id,
        "practice_id": practice_id,
        "name": name,
        "filename": filename,
        "file_path": str(file_path),
        "created_at": datetime.utcnow().isoformat(),
        "version": 1,
        "config": DEFAULT_TEMPLATE_CONFIG,
        "normalisation": normalisation,
    }
    _write_json(template_dir / "info.json", template)

    practice = get_practice(practice_id)
    if practice and not practice.get("default_template_id"):
        practice["default_template_id"] = template_id
        save_practice(practice)
    return template


def get_template(practice_id: str, template_id: str) -> dict | None:
    return _read_json(PRACTICES_DIR / practice_id / "templates" / template_id / "info.json")


def save_template(template: dict) -> None:
    template["version"] = template.get("version", 1) + 1
    _write_json(PRACTICES_DIR / template["practice_id"] / "templates" / template["id"] / "info.json", template)


def list_templates(practice_id: str) -> list[dict]:
    templates_dir = PRACTICES_DIR / practice_id / "templates"
    if not templates_dir.exists():
        return []
    out = []
    for d in sorted(templates_dir.iterdir()):
        info = _read_json(d / "info.json")
        if info:
            out.append(info)
    return out


# ---------- clients (scoped to a practice) ----------

def create_client(practice_id: str, name: str, template_id: str | None = None) -> dict:
    client_id = _new_id("client")
    client = {
        "id": client_id, "practice_id": practice_id, "name": name,
        "template_id": template_id, "created_at": datetime.utcnow().isoformat(),
    }
    _write_json(CLIENTS_DIR / client_id / "info.json", client)
    return client


def get_client(client_id: str) -> dict | None:
    return _read_json(CLIENTS_DIR / client_id / "info.json")


def list_clients(practice_id: str | None = None) -> list[dict]:
    if not CLIENTS_DIR.exists():
        return []
    out = []
    for d in sorted(CLIENTS_DIR.iterdir()):
        info = _read_json(d / "info.json")
        if info and (practice_id is None or info.get("practice_id") == practice_id):
            out.append(info)
    return out


# ---------- jobs ----------

def create_job(
    client_id: str, client_name: str,
    current_period_start: str, current_period_end: str,
    comparative_period_start: str | None, comparative_period_end: str | None,
) -> dict:
    job_id = _new_id("job")
    job = {
        "id": job_id,
        "client_id": client_id,
        "client_name": client_name,
        "current_period_start": current_period_start,
        "current_period_end": current_period_end,
        "comparative_period_start": comparative_period_start,
        "comparative_period_end": comparative_period_end,
        "current_label": _period_label(current_period_start, current_period_end),
        "comparative_label": _period_label(comparative_period_start, comparative_period_end) if comparative_period_end else "",
        "created_at": datetime.utcnow().isoformat(),
        "status": "draft",
        "uploads": {},
    }
    _write_json(JOBS_DIR / job_id / "job.json", job)
    return job


def _period_label(start: str, end: str) -> str:
    start_d, end_d = datetime.fromisoformat(start), datetime.fromisoformat(end)
    if (start_d.month, start_d.day) == (1, 1) and (end_d.month, end_d.day) == (12, 31):
        return f"Year ended {end_d:%d %B %Y}"
    return f"{start_d:%d %b %Y} to {end_d:%d %b %Y}"


def get_job(job_id: str) -> dict | None:
    return _read_json(JOBS_DIR / job_id / "job.json")


def save_job(job: dict) -> None:
    _write_json(JOBS_DIR / job["id"] / "job.json", job)


def list_jobs(client_id: str | None = None) -> list[dict]:
    if not JOBS_DIR.exists():
        return []
    out = []
    for d in sorted(JOBS_DIR.iterdir(), reverse=True):
        job = _read_json(d / "job.json")
        if job and (client_id is None or job["client_id"] == client_id):
            out.append(job)
    return out


def job_dir(job_id: str) -> Path:
    return JOBS_DIR / job_id


def uploads_dir(job_id: str) -> Path:
    d = JOBS_DIR / job_id / "uploads"
    d.mkdir(parents=True, exist_ok=True)
    return d


def output_dir(job_id: str) -> Path:
    d = JOBS_DIR / job_id / "output"
    d.mkdir(parents=True, exist_ok=True)
    return d


def add_upload(job: dict, report_type: str, period: str, platform: str, filename: str, saved_path: str, columns: list[str]) -> str:
    upload_id = _new_id("upload")
    job["uploads"][upload_id] = {
        "id": upload_id,
        "report_type": report_type,
        "period": period,
        "platform": platform,
        "filename": filename,
        "path": saved_path,
        "columns": columns,
        "mapping": None,
        "confirmed": False,
    }
    save_job(job)
    return upload_id
