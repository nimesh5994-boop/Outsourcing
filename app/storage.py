"""Postgres-backed storage for practices, templates, clients, jobs, and the
files that belong to them (uploaded reports, template .xlsx files,
generated working papers).

Started out as JSON files on the local filesystem - fine for a single
long-running server, but broken on a serverless platform (Vercel): each
invocation can land on a different instance with its own ephemeral
filesystem, so a file written during upload might not be there when
generate() runs a moment later. Postgres (Neon) fixes that by making state
shared and durable across invocations; small/medium files (uploaded
reports, template workbooks, generated packs) are stored as BYTEA in the
same database rather than adding a second storage service - simplest
architecture that actually works on Vercel, with a documented upgrade path
to real object storage (Vercel Blob / S3) if file sizes grow enough that
BYTEA and response-size limits become a real constraint (see README).

Two generic tables back nearly everything:
  entities(kind, id, parent_id, data JSONB) - practices/templates/clients/jobs,
    each just a dict, the same shape they were as JSON files.
  files(id, job_id, kind, filename, content BYTEA) - the raw bytes for an
    upload, a template, or a generated output workbook.
mapping_profiles is its own small table (composite key, not a blob) since
it's looked up by (client, report_type, platform), not by id.
"""
import os
import uuid
from datetime import datetime

import psycopg
from psycopg.types.json import Jsonb

# Default customization config for a newly-uploaded template. A practice
# edits this to control which schedules get generated, where each one
# lands in the template, and the thresholds/cell conventions used - so
# adding a second (or fifth) template format later is configuration, not
# new code.
DEFAULT_TEMPLATE_CONFIG = {
    "schedules": {
        "index": {"enabled": True, "insert_after_sheet": None},  # None = ahead of the template's own sheets
        "tb_lead_schedule": {"enabled": True, "insert_after_sheet": None},
        "tb_balance_check": {"enabled": True, "insert_after_sheet": None},
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
        "contact_coding_consistency": {"enabled": True, "insert_after_sheet": None},
        "duplicate_check": {"enabled": True, "insert_after_sheet": None},
        "unusual_posting_dates": {"enabled": True, "insert_after_sheet": None},
        "dla_review": {"enabled": True, "insert_after_sheet": None},
        "dividend_reserves_review": {"enabled": True, "insert_after_sheet": None},
        "petty_cash_review": {"enabled": True, "insert_after_sheet": None},
        "loan_facility_review": {"enabled": True, "insert_after_sheet": None},
        "compliance_checklist": {"enabled": True, "insert_after_sheet": None},
        "points_forward": {"enabled": True, "insert_after_sheet": None},
    },
    "header_cells": {"client_name_cell": "A1", "period_cell": "A2", "schedule_title_cell": "A3"},
    "materiality": {"default_amount": 500, "variance_pct_threshold": 0.10},
    "numbering": {"style": "sequential", "start_at": 1},
}

SCHEMA_STATEMENTS = [
    """CREATE TABLE IF NOT EXISTS entities (
        kind TEXT NOT NULL,
        id TEXT NOT NULL,
        parent_id TEXT,
        data JSONB NOT NULL,
        created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
        PRIMARY KEY (kind, id)
    )""",
    "CREATE INDEX IF NOT EXISTS idx_entities_kind_parent ON entities (kind, parent_id)",
    """CREATE TABLE IF NOT EXISTS files (
        id TEXT PRIMARY KEY,
        job_id TEXT,
        kind TEXT NOT NULL,
        filename TEXT NOT NULL,
        content BYTEA NOT NULL,
        content_type TEXT,
        created_at TIMESTAMPTZ NOT NULL DEFAULT now()
    )""",
    "CREATE INDEX IF NOT EXISTS idx_files_job ON files (job_id)",
    """CREATE TABLE IF NOT EXISTS mapping_profiles (
        client_id TEXT NOT NULL,
        report_type TEXT NOT NULL,
        platform TEXT NOT NULL,
        mapping JSONB NOT NULL,
        PRIMARY KEY (client_id, report_type, platform)
    )""",
    """CREATE TABLE IF NOT EXISTS users (
        id TEXT PRIMARY KEY,
        practice_id TEXT NOT NULL,
        email TEXT NOT NULL,
        password_hash TEXT NOT NULL,
        name TEXT NOT NULL,
        role TEXT NOT NULL,
        created_at TIMESTAMPTZ NOT NULL DEFAULT now()
    )""",
    "CREATE UNIQUE INDEX IF NOT EXISTS idx_users_email ON users (lower(email))",
    "CREATE INDEX IF NOT EXISTS idx_users_practice ON users (practice_id)",
    # which specific clients a 'preparer' can see - irrelevant for
    # partner/manager, who see every client in their own practice.
    """CREATE TABLE IF NOT EXISTS client_access (
        user_id TEXT NOT NULL,
        client_id TEXT NOT NULL,
        PRIMARY KEY (user_id, client_id)
    )""",
]

_conn: psycopg.Connection | None = None


def _connect() -> psycopg.Connection:
    dsn = os.environ["DATABASE_URL"]
    conn = psycopg.connect(dsn, autocommit=True)
    with conn.cursor() as cur:
        for stmt in SCHEMA_STATEMENTS:
            cur.execute(stmt)
    return conn


def _get_conn() -> psycopg.Connection:
    """Reuses one connection for the life of the process (a serverless
    instance can be reused across several invocations) - reconnects
    transparently if it's been closed or dropped (idle timeout, cold
    start)."""
    global _conn
    if _conn is None or _conn.closed:
        _conn = _connect()
    return _conn


def _new_id(prefix: str) -> str:
    return f"{prefix}_{uuid.uuid4().hex[:10]}"


# ---------- generic entity storage (practices/templates/clients/jobs) ----------

def _get_entity(kind: str, entity_id: str) -> dict | None:
    with _get_conn().cursor() as cur:
        cur.execute("SELECT data FROM entities WHERE kind = %s AND id = %s", (kind, entity_id))
        row = cur.fetchone()
        return row[0] if row else None


def _put_entity(kind: str, entity_id: str, parent_id: str | None, data: dict) -> None:
    with _get_conn().cursor() as cur:
        cur.execute(
            """INSERT INTO entities (kind, id, parent_id, data) VALUES (%s, %s, %s, %s)
               ON CONFLICT (kind, id) DO UPDATE SET data = EXCLUDED.data, parent_id = EXCLUDED.parent_id""",
            (kind, entity_id, parent_id, Jsonb(data)),
        )


def _list_entities(kind: str, parent_id: str | None = None, order_by_id_desc: bool = False) -> list[dict]:
    order = "id DESC" if order_by_id_desc else "id ASC"
    with _get_conn().cursor() as cur:
        if parent_id is None:
            cur.execute(f"SELECT data FROM entities WHERE kind = %s ORDER BY {order}", (kind,))
        else:
            cur.execute(f"SELECT data FROM entities WHERE kind = %s AND parent_id = %s ORDER BY {order}", (kind, parent_id))
        return [row[0] for row in cur.fetchall()]


# ---------- file storage (uploads, templates, generated output) ----------

def save_file(kind: str, job_id: str | None, filename: str, content: bytes, content_type: str | None = None) -> str:
    file_id = _new_id("file")
    with _get_conn().cursor() as cur:
        cur.execute(
            "INSERT INTO files (id, job_id, kind, filename, content, content_type) VALUES (%s, %s, %s, %s, %s, %s)",
            (file_id, job_id, kind, filename, content, content_type),
        )
    return file_id


def load_file(file_id: str) -> bytes | None:
    with _get_conn().cursor() as cur:
        cur.execute("SELECT content FROM files WHERE id = %s", (file_id,))
        row = cur.fetchone()
        return bytes(row[0]) if row else None


# ---------- practices ----------

def create_practice(name: str) -> dict:
    practice_id = _new_id("practice")
    practice = {"id": practice_id, "name": name, "created_at": datetime.utcnow().isoformat(), "default_template_id": None}
    _put_entity("practice", practice_id, None, practice)
    return practice


def get_practice(practice_id: str) -> dict | None:
    return _get_entity("practice", practice_id)


def save_practice(practice: dict) -> None:
    _put_entity("practice", practice["id"], None, practice)


def list_practices() -> list[dict]:
    return _list_entities("practice")


# ---------- templates (scoped to a practice) ----------

def _normalise_template_file(content: bytes) -> tuple[bytes, dict]:
    """One-time cleanup at upload time (not repeated per job): loading and
    re-saving through openpyxl drops calcChain/printer-settings bloat and
    shrinks the file substantially - confirmed on a real 63-sheet template,
    24MB down to 11MB with cell data, formulas, and formatting all intact.
    It also drops embedded images and dropdown data-validation lists, which
    is the deliberate trade-off: doing this once at setup keeps every later
    job generation fast, at the cost of re-adding a logo/validation lists
    once per template rather than never. Falls back to the original bytes
    untouched if the file can't be parsed."""
    import io

    import openpyxl

    original_size = len(content)
    try:
        wb = openpyxl.load_workbook(io.BytesIO(content))
        buf = io.BytesIO()
        wb.save(buf)
        normalised_content = buf.getvalue()
        normalised = True
    except Exception:
        normalised_content = content
        normalised = False
    return normalised_content, {
        "normalised": normalised,
        "original_size_bytes": original_size,
        "stored_size_bytes": len(normalised_content),
    }


def create_template(practice_id: str, name: str, file_bytes: bytes, filename: str) -> dict:
    template_id = _new_id("template")
    normalised_content, normalisation = _normalise_template_file(file_bytes)
    file_id = save_file("template", None, filename, normalised_content)

    template = {
        "id": template_id,
        "practice_id": practice_id,
        "name": name,
        "filename": filename,
        "file_id": file_id,
        "created_at": datetime.utcnow().isoformat(),
        "version": 1,
        "config": DEFAULT_TEMPLATE_CONFIG,
        "normalisation": normalisation,
    }
    _put_entity("template", template_id, practice_id, template)

    practice = get_practice(practice_id)
    if practice and not practice.get("default_template_id"):
        practice["default_template_id"] = template_id
        save_practice(practice)
    return template


def get_template(practice_id: str, template_id: str) -> dict | None:
    template = _get_entity("template", template_id)
    return template if template and template.get("practice_id") == practice_id else None


def save_template(template: dict) -> None:
    template["version"] = template.get("version", 1) + 1
    _put_entity("template", template["id"], template["practice_id"], template)


def list_templates(practice_id: str) -> list[dict]:
    return _list_entities("template", practice_id)


# ---------- clients (scoped to a practice) ----------

def create_client(practice_id: str, name: str, template_id: str | None = None) -> dict:
    client_id = _new_id("client")
    client = {
        "id": client_id, "practice_id": practice_id, "name": name,
        "template_id": template_id, "created_at": datetime.utcnow().isoformat(),
        # {report_type: note text} - how this client's exports of that type
        # should be read/treated (e.g. "VAT export has a Detail tab, use it
        # to reconcile against nominal activity"). Persists per client
        # across jobs, and is shown alongside that section's uploads.
        "report_notes": {},
    }
    _put_entity("client", client_id, practice_id, client)
    return client


def get_client(client_id: str) -> dict | None:
    return _get_entity("client", client_id)


def save_client(client: dict) -> None:
    _put_entity("client", client["id"], client["practice_id"], client)


def list_clients(practice_id: str | None = None) -> list[dict]:
    return _list_entities("client", practice_id)


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
    _put_entity("job", job_id, client_id, job)
    return job


def _period_label(start: str, end: str) -> str:
    start_d, end_d = datetime.fromisoformat(start), datetime.fromisoformat(end)
    if (start_d.month, start_d.day) == (1, 1) and (end_d.month, end_d.day) == (12, 31):
        return f"Year ended {end_d:%d %B %Y}"
    return f"{start_d:%d %b %Y} to {end_d:%d %b %Y}"


def get_job(job_id: str) -> dict | None:
    return _get_entity("job", job_id)


def save_job(job: dict) -> None:
    _put_entity("job", job["id"], job["client_id"], job)


def list_jobs(client_id: str | None = None) -> list[dict]:
    return _list_entities("job", client_id, order_by_id_desc=True)


def add_upload(job: dict, report_type: str, period: str, platform: str, filename: str, content: bytes, columns: list[str]) -> str:
    file_id = save_file("upload", job["id"], filename, content)
    upload_id = _new_id("upload")
    job["uploads"][upload_id] = {
        "id": upload_id,
        "report_type": report_type,
        "period": period,
        "platform": platform,
        "filename": filename,
        "file_id": file_id,
        "columns": columns,
        "mapping": None,
        "confirmed": False,
    }
    save_job(job)
    return upload_id


# ---------- reusable client mapping profiles ----------

def load_mapping_profile(client_id: str, report_type: str, platform: str) -> dict | None:
    with _get_conn().cursor() as cur:
        cur.execute(
            "SELECT mapping FROM mapping_profiles WHERE client_id = %s AND report_type = %s AND platform = %s",
            (client_id, report_type, platform),
        )
        row = cur.fetchone()
        return row[0] if row else None


def save_mapping_profile(client_id: str, report_type: str, platform: str, mapping: dict) -> None:
    with _get_conn().cursor() as cur:
        cur.execute(
            """INSERT INTO mapping_profiles (client_id, report_type, platform, mapping) VALUES (%s, %s, %s, %s)
               ON CONFLICT (client_id, report_type, platform) DO UPDATE SET mapping = EXCLUDED.mapping""",
            (client_id, report_type, platform, Jsonb(mapping)),
        )


# ---------- users (each belongs to exactly one practice) + client access ----------
# Roles: partner (full access within the practice, incl. managing users and
# templates), manager (everything except managing users), preparer (only the
# clients explicitly granted below). A real table rather than the generic
# entities blob, since email needs a case-insensitive unique index and
# password_hash should never round-trip through a generic JSONB dict by
# accident.

def _user_row_to_dict(row) -> dict:
    return {
        "id": row[0], "practice_id": row[1], "email": row[2], "password_hash": row[3],
        "name": row[4], "role": row[5],
        "created_at": row[6].isoformat() if hasattr(row[6], "isoformat") else row[6],
    }


_USER_COLUMNS = "id, practice_id, email, password_hash, name, role, created_at"


def create_user(practice_id: str, email: str, password_hash: str, name: str, role: str) -> dict:
    user_id = _new_id("user")
    email = email.strip().lower()
    with _get_conn().cursor() as cur:
        cur.execute(
            "INSERT INTO users (id, practice_id, email, password_hash, name, role) VALUES (%s, %s, %s, %s, %s, %s)",
            (user_id, practice_id, email, password_hash, name, role),
        )
    return get_user(user_id)


def get_user(user_id: str) -> dict | None:
    with _get_conn().cursor() as cur:
        cur.execute(f"SELECT {_USER_COLUMNS} FROM users WHERE id = %s", (user_id,))
        row = cur.fetchone()
        return _user_row_to_dict(row) if row else None


def get_user_by_email(email: str) -> dict | None:
    with _get_conn().cursor() as cur:
        cur.execute(f"SELECT {_USER_COLUMNS} FROM users WHERE lower(email) = lower(%s)", (email.strip(),))
        row = cur.fetchone()
        return _user_row_to_dict(row) if row else None


def list_users(practice_id: str) -> list[dict]:
    with _get_conn().cursor() as cur:
        cur.execute(f"SELECT {_USER_COLUMNS} FROM users WHERE practice_id = %s ORDER BY created_at ASC", (practice_id,))
        return [_user_row_to_dict(row) for row in cur.fetchall()]


def count_users() -> int:
    with _get_conn().cursor() as cur:
        cur.execute("SELECT count(*) FROM users")
        return cur.fetchone()[0]


def delete_user(user_id: str) -> None:
    with _get_conn().cursor() as cur:
        cur.execute("DELETE FROM client_access WHERE user_id = %s", (user_id,))
        cur.execute("DELETE FROM users WHERE id = %s", (user_id,))


def set_client_access(user_id: str, client_ids: list[str]) -> None:
    """Replaces the full set of clients a preparer can see - simplest correct
    model for a checkbox-list UI (submit the whole selection, not deltas)."""
    with _get_conn().cursor() as cur:
        cur.execute("DELETE FROM client_access WHERE user_id = %s", (user_id,))
        for client_id in client_ids:
            cur.execute(
                "INSERT INTO client_access (user_id, client_id) VALUES (%s, %s) ON CONFLICT DO NOTHING",
                (user_id, client_id),
            )


def has_client_access(user_id: str, client_id: str) -> bool:
    with _get_conn().cursor() as cur:
        cur.execute("SELECT 1 FROM client_access WHERE user_id = %s AND client_id = %s", (user_id, client_id))
        return cur.fetchone() is not None


def list_client_access(user_id: str) -> list[str]:
    with _get_conn().cursor() as cur:
        cur.execute("SELECT client_id FROM client_access WHERE user_id = %s", (user_id,))
        return [row[0] for row in cur.fetchall()]
