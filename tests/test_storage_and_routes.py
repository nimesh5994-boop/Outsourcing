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
import io
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
    monkeypatch.setenv("SECRET_KEY", "test-secret-key-not-for-production")

    from app import storage
    storage._conn = None  # force reconnect against the schema-scoped DSN

    from app.main import app
    with TestClient(app) as c:
        yield c

    storage._conn = None
    admin_conn = psycopg.connect(TEST_DATABASE_URL, autocommit=True)
    admin_conn.execute(f'DROP SCHEMA "{schema}" CASCADE')
    admin_conn.close()


def _signup(c: TestClient, practice_name="Acme & Co", admin_name="Alex Partner",
            admin_email="alex@acme.test", admin_password="hunter2hunter") -> str:
    """Creates a practice + its first (partner) user and logs the client in
    via the session cookie, exactly as a real signup would. Returns the new
    practice_id."""
    resp = c.post(
        "/practices",
        data={
            "practice_name": practice_name, "admin_name": admin_name,
            "admin_email": admin_email, "admin_password": admin_password,
        },
        follow_redirects=False,
    )
    assert resp.status_code == 303
    return resp.headers["location"].rsplit("/", 1)[-1]


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
    """signup (practice + partner login) -> template upload (normalised) ->
    client -> job -> report upload -> generate -> download, all over real
    HTTP requests against a real Postgres-backed app instance, as the
    logged-in partner."""
    c = http_client
    practice_id = _signup(c)

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
            files={"files": (tb_path.name, fh,
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            follow_redirects=False,
        )
    assert resp.status_code == 303  # Xero-native TB auto-confirms straight through, no mapping-confirm redirect

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


def test_unauthenticated_request_redirects_to_login(http_client):
    c = http_client
    resp = c.get("/practices/practice_doesnotmatter", follow_redirects=False)
    assert resp.status_code == 303
    assert resp.headers["location"].startswith("/login")


def test_wrong_password_rejected(http_client):
    c = http_client
    _signup(c, admin_email="owner@acme.test", admin_password="correct-horse-battery")
    resp = c.post("/login", data={"email": "owner@acme.test", "password": "wrong"}, follow_redirects=False)
    assert resp.status_code == 400
    assert "Invalid email or password" in resp.text


def test_preparer_scoped_to_granted_clients_only(http_client):
    """The core RBAC guarantee: a preparer can see/act on a client they've
    been granted access to, and gets a 403 - not the client's data - for
    one they haven't, even though both clients are in the same practice."""
    c = http_client
    practice_id = _signup(c, admin_email="partner@acme.test")

    resp = c.post(f"/practices/{practice_id}/clients", data={"name": "Client A"}, follow_redirects=False)
    client_a_id = resp.headers["location"].rsplit("/", 1)[-1]
    resp = c.post(f"/practices/{practice_id}/clients", data={"name": "Client B"}, follow_redirects=False)
    client_b_id = resp.headers["location"].rsplit("/", 1)[-1]

    resp = c.post(f"/practices/{practice_id}/users", data={
        "name": "Prep One", "email": "prep@acme.test", "password": "prepper-pass", "role": "preparer",
    }, follow_redirects=False)
    assert resp.status_code == 303

    from app import storage
    prep_user = storage.get_user_by_email("prep@acme.test")
    resp = c.post(
        f"/practices/{practice_id}/users/{prep_user['id']}/client-access",
        data={"client_ids": [client_a_id]}, follow_redirects=False,
    )
    assert resp.status_code == 303

    # log out the partner, log in as the preparer
    c.post("/logout")
    resp = c.post("/login", data={"email": "prep@acme.test", "password": "prepper-pass"}, follow_redirects=False)
    assert resp.status_code == 303

    assert c.get(f"/clients/{client_a_id}").status_code == 200
    resp = c.get(f"/clients/{client_b_id}", follow_redirects=False)
    assert resp.status_code == 403

    # preparer's client list only shows the granted client
    resp = c.get(f"/practices/{practice_id}/clients")
    assert "Client A" in resp.text
    assert "Client B" not in resp.text

    # preparer role restrictions: no creating clients, no managing users/templates
    assert c.post(f"/practices/{practice_id}/clients", data={"name": "Client C"}, follow_redirects=False).status_code == 403
    assert c.get(f"/practices/{practice_id}/users", follow_redirects=False).status_code == 403
    assert c.post(f"/practices/{practice_id}/templates", data={"name": "X"}, follow_redirects=False).status_code == 403


def test_manager_sees_all_clients_but_cannot_manage_users(http_client):
    c = http_client
    practice_id = _signup(c, admin_email="partner2@acme.test")
    c.post(f"/practices/{practice_id}/clients", data={"name": "Only Client"}, follow_redirects=False)

    c.post(f"/practices/{practice_id}/users", data={
        "name": "Mgr One", "email": "mgr@acme.test", "password": "manager-pass", "role": "manager",
    }, follow_redirects=False)

    c.post("/logout")
    c.post("/login", data={"email": "mgr@acme.test", "password": "manager-pass"}, follow_redirects=False)

    resp = c.get(f"/practices/{practice_id}/clients")
    assert "Only Client" in resp.text  # managers see every client, no grant needed

    assert c.get(f"/practices/{practice_id}/users", follow_redirects=False).status_code == 403


def test_cross_practice_access_denied(http_client):
    """A user from one practice can't reach another practice's pages, even
    though both exist in the same database."""
    c = http_client
    practice_a_id = _signup(c, admin_email="a@firm-a.test")

    from app.main import app as fastapi_app
    other = TestClient(fastapi_app)
    resp = other.post("/practices", data={
        "practice_name": "Firm B", "admin_name": "B Partner",
        "admin_email": "b@firm-b.test", "admin_password": "firm-b-password",
    }, follow_redirects=False)
    assert resp.status_code == 303

    resp = other.get(f"/practices/{practice_a_id}", follow_redirects=False)
    assert resp.status_code == 404


def _make_pdf_trial_balance() -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

    rows = [
        ["Account Code", "Account Name", "Debit", "Credit"],
        ["1000", "Bank Current Account", "12000.00", ""],
        ["4000", "Sales", "", "75000.00"],
    ]
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4)
    table = Table(rows)
    table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black)]))
    doc.build([table])
    return buf.getvalue()


def test_bulk_upload_mixed_files_auto_detect_and_confirm_chain(http_client):
    """The real workflow this feature exists for: drop in several files of
    different kinds and formats at once (a genuine Xero export, a generic
    CSV, and a PDF) with no report-type/platform/period picked up front.
    The Xero file should auto-confirm silently (real parser validation, not
    a guess); the other two should get queued for a quick confirm each,
    chained one after another; and the PDF (uploaded alongside a current-
    period TB) should be guessed as the comparative TB via the "second
    upload of this type" heuristic, since it has no period text of its own."""
    pytest.importorskip("reportlab", reason="reportlab is a dev-only dependency for building test PDFs")
    c = http_client
    practice_id = _signup(c, admin_email="bulk@acme.test")

    resp = c.post(f"/practices/{practice_id}/clients", data={"name": "Bulk Client"}, follow_redirects=False)
    client_id = resp.headers["location"].rsplit("/", 1)[-1]
    resp = c.post(f"/clients/{client_id}/jobs", data={
        "current_period_start": "2025-01-01", "current_period_end": "2025-12-31",
        "comparative_period_start": "2024-01-01", "comparative_period_end": "2024-12-31",
    }, follow_redirects=False)
    job_id = resp.headers["location"].rsplit("/", 1)[-1]

    tb_path = SAMPLE_DIR / "trial_balance_current_xero.xlsx"
    bank_path = SAMPLE_DIR / "bank_statement_current.csv"
    files_payload = [
        ("files", (tb_path.name, tb_path.read_bytes(),
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")),
        ("files", (bank_path.name, bank_path.read_bytes(), "text/csv")),
        ("files", ("prior_year_tb.pdf", _make_pdf_trial_balance(), "application/pdf")),
    ]
    resp = c.post(f"/jobs/{job_id}/uploads", files=files_payload, follow_redirects=False)
    assert resp.status_code == 303

    from app import storage
    job = storage.get_job(job_id)
    by_filename = {u["filename"]: u for u in job["uploads"].values()}
    assert by_filename[tb_path.name]["confirmed"] is True  # Xero-native: no confirm step needed
    assert by_filename[tb_path.name]["report_type"] == "trial_balance"
    assert by_filename[bank_path.name]["report_type"] == "bank_statement"
    assert by_filename["prior_year_tb.pdf"]["report_type"] == "trial_balance"

    # walk the confirm chain the redirect points at, for every unconfirmed upload
    location = resp.headers["location"]
    hops = 0
    while location and "/mapping" in location and hops < 5:
        hops += 1
        assert c.get(location).status_code == 200
        resp = c.post(location, data={"action": "confirm"}, follow_redirects=False)
        assert resp.status_code == 303
        location = resp.headers["location"]
    assert hops == 2  # bank statement + PDF both needed a confirm; the Xero file didn't

    job = storage.get_job(job_id)
    assert all(u["confirmed"] for u in job["uploads"].values())
    by_filename = {u["filename"]: u for u in job["uploads"].values()}
    assert by_filename["prior_year_tb.pdf"]["period"] == "comparative"

    resp = c.post(f"/jobs/{job_id}/generate", follow_redirects=False)
    assert resp.status_code == 303
    resp = c.get(f"/jobs/{job_id}/download")
    assert resp.status_code == 200
    assert len(resp.content) > 1000


def test_upload_route_rejects_confirm_with_no_report_type(http_client):
    """Guard against silently confirming an upload the system couldn't
    classify and the user didn't pick a type for either - that would mark
    it 'confirmed' with an empty mapping, contributing nothing to the
    generated workbook with no visible sign anything was wrong."""
    c = http_client
    practice_id = _signup(c, admin_email="guard@acme.test")
    resp = c.post(f"/practices/{practice_id}/clients", data={"name": "Guard Client"}, follow_redirects=False)
    client_id = resp.headers["location"].rsplit("/", 1)[-1]
    resp = c.post(f"/clients/{client_id}/jobs", data={
        "current_period_start": "2025-01-01", "current_period_end": "2025-12-31",
    }, follow_redirects=False)
    job_id = resp.headers["location"].rsplit("/", 1)[-1]

    resp = c.post(
        f"/jobs/{job_id}/uploads",
        files={"files": ("mystery.csv", b"Foo,Bar\n1,2\n", "text/csv")},
        follow_redirects=False,
    )
    location = resp.headers["location"]
    assert "/mapping" in location

    from app import storage
    job = storage.get_job(job_id)
    upload = next(iter(job["uploads"].values()))
    assert upload["report_type"] == ""

    resp = c.post(location, data={"action": "confirm"}, follow_redirects=False)
    assert resp.status_code == 303
    assert resp.headers["location"] == location  # bounced back, not confirmed

    job = storage.get_job(job_id)
    upload = next(iter(job["uploads"].values()))
    assert upload["confirmed"] is False
