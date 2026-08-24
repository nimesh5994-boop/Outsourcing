"""Integration tests for the Postgres-backed storage layer and the FastAPI
HTTP routes that depend on it - the part of the storage.py rewrite (JSON
files on disk -> Postgres entities/files/mapping_profiles tables) that the
rest of the suite never exercises, since every other test drives the
computation modules directly with in-memory DataFrames/fixtures.

Runs against a real, throwaway Postgres schema (not the app's production
Neon database) so the whole path - entity CRUD, file BYTEA storage,
template normalisation, and the practice -> template -> client -> job ->
upload -> generate -> download HTTP flow - is proven against a real
database engine, not a mock. Skipped automatically when no test database
is reachable (e.g. a machine without Postgres installed); set
TEST_DATABASE_URL to point at one explicitly.
"""
import os
import uuid
from pathlib import Path

import psycopg
import pytest
from fastapi.testclient import TestClient

SAMPLE_DIR = Path(__file__).resolve().parent.parent / "sample_data"

TEST_DATABASE_URL = os.environ.get(
    "TEST_DATABASE_URL", "postgresql://wpa_test:wpa_test@127.0.0.1:5432/wpa_test"
)


def _database_available() -> bool:
    try:
        conn = psycopg.connect(TEST_DATABASE_URL, connect_timeout=3)
        conn.close()
        return True
    except Exception:
        return False


pytestmark = pytest.mark.skipif(
    not _database_available(),
    reason=(
        f"no reachable Postgres test database at {TEST_DATABASE_URL!r} - "
        "set TEST_DATABASE_URL, or run one locally, to exercise storage.py"
    ),
)


@pytest.fixture
def http_client(monkeypatch):
    """A TestClient wired to a throwaway schema in the test database, so
    repeated/parallel test runs never collide and nothing here can ever
    touch the real app's data."""
    schema = f"test_{uuid.uuid4().hex[:12]}"
    admin_conn = psycopg.connect(TEST_DATABASE_URL, autocommit=True)
    admin_conn.execute(f'CREATE SCHEMA "{schema}"')
    admin_conn.close()

    separator = "&" if "?" in TEST_DATABASE_URL else "?"
    scoped_dsn = f"{TEST_DATABASE_URL}{separator}options=-csearch_path%3D{schema}"
    monkeypatch.setenv("DATABASE_URL", scoped_dsn)

    from app import storage
    storage._conn = None  # force reconnect against the schema-scoped DSN

    from app.main import app
    with TestClient(app) as c:
        yield c

    storage._conn = None
    admin_conn = psycopg.connect(TEST_DATABASE_URL, autocommit=True)
    admin_conn.execute(f'DROP SCHEMA "{schema}" CASCADE')
    admin_conn.close()


def _make_template_bytes() -> bytes:
    import io

    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cover Page"
    ws["A1"] = "FIRM TEMPLATE COVER"
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def test_storage_entity_and_file_roundtrip(http_client):
    """Direct storage.py calls: entities and file BYTEA both survive a
    reconnect-free roundtrip through the real database."""
    from app import storage

    practice = storage.create_practice("Test Practice")
    assert practice["id"]
    assert storage.get_practice(practice["id"])["name"] == "Test Practice"

    file_id = storage.save_file("upload", None, "sample.txt", b"hello world", content_type="text/plain")
    assert storage.load_file(file_id) == b"hello world"

    storage.save_mapping_profile("client_x", "trial_balance", "xero", {"Account": "account_name"})
    assert storage.load_mapping_profile("client_x", "trial_balance", "xero") == {"Account": "account_name"}
    assert storage.load_mapping_profile("client_x", "trial_balance", "sage") is None


def test_full_http_flow_practice_to_download(http_client):
    """practice -> template upload (normalised) -> client -> job -> report
    upload -> generate -> download, all over real HTTP requests against a
    real Postgres-backed app instance."""
    c = http_client

    resp = c.post("/practices", data={"name": "Acme & Co"}, follow_redirects=False)
    assert resp.status_code == 303
    practice_id = resp.headers["location"].rsplit("/", 1)[-1]

    resp = c.post(
        f"/practices/{practice_id}/templates",
        data={"name": "Standard Template"},
        files={"file": ("template.xlsx", _make_template_bytes(),
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        follow_redirects=False,
    )
    assert resp.status_code == 303

    from app import storage
    templates_list = storage.list_templates(practice_id)
    assert len(templates_list) == 1
    template_id = templates_list[0]["id"]
    assert templates_list[0]["normalisation"]["normalised"] is True

    resp = c.post(
        f"/practices/{practice_id}/clients",
        data={"name": "Sample Client Ltd", "template_id": template_id},
        follow_redirects=False,
    )
    assert resp.status_code == 303
    client_id = resp.headers["location"].rsplit("/", 1)[-1]

    resp = c.post(
        f"/clients/{client_id}/jobs",
        data={"current_period_start": "2024-01-01", "current_period_end": "2024-12-31"},
        follow_redirects=False,
    )
    assert resp.status_code == 303
    job_id = resp.headers["location"].rsplit("/", 1)[-1]

    tb_path = SAMPLE_DIR / "trial_balance_current_xero.xlsx"
    with open(tb_path, "rb") as fh:
        resp = c.post(
            f"/jobs/{job_id}/uploads",
            data={"report_type": "trial_balance", "period": "current", "platform": "xero"},
            files={"file": (tb_path.name, fh,
                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            follow_redirects=False,
        )
    assert resp.status_code == 303

    job = storage.get_job(job_id)
    assert len(job["uploads"]) == 1
    upload = next(iter(job["uploads"].values()))
    assert storage.load_file(upload["file_id"]) is not None

    resp = c.post(f"/jobs/{job_id}/generate", follow_redirects=False)
    assert resp.status_code == 303

    resp = c.get(f"/jobs/{job_id}/download")
    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert len(resp.content) > 1000

    import io

    import openpyxl

    wb = openpyxl.load_workbook(io.BytesIO(resp.content))
    assert "Cover Page" in wb.sheetnames  # the template's own sheet survived untouched
    assert wb["Cover Page"]["A1"].value == "FIRM TEMPLATE COVER"
