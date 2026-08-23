"""Filesystem-backed storage for clients and jobs.

Deliberately simple (JSON files on disk) for an internal single-instance
tool. Swapping this for a real database later is a contained change since
every other module only talks to the functions in this file.
"""
import json
import uuid
from datetime import datetime
from pathlib import Path

DATA_DIR = Path("data")
CLIENTS_DIR = DATA_DIR / "clients"
JOBS_DIR = DATA_DIR / "jobs"


def _new_id(prefix: str) -> str:
    return f"{prefix}_{uuid.uuid4().hex[:10]}"


def _read_json(path: Path) -> dict | None:
    if path.exists():
        return json.loads(path.read_text())
    return None


def _write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, default=str))


# ---------- clients ----------

def create_client(name: str) -> dict:
    client_id = _new_id("client")
    client = {"id": client_id, "name": name, "created_at": datetime.utcnow().isoformat()}
    _write_json(CLIENTS_DIR / client_id / "info.json", client)
    return client


def get_client(client_id: str) -> dict | None:
    return _read_json(CLIENTS_DIR / client_id / "info.json")


def list_clients() -> list[dict]:
    if not CLIENTS_DIR.exists():
        return []
    out = []
    for d in sorted(CLIENTS_DIR.iterdir()):
        info = _read_json(d / "info.json")
        if info:
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
