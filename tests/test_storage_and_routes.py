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


def _make_multisheet_vat_workbook() -> bytes:
    from openpyxl import Workbook

    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    summary.append(["Box 1", "Box 2", "Box 3", "Box 4", "Box 5", "Box 6", "Box 7", "Box 8", "Box 9"])
    summary.append([1000, 0, 1000, 200, 800, 15000, 5000, 0, 0])
    detail = wb.create_sheet("Detail")
    detail.append(["Date", "Account Code", "Description", "Debit", "Credit"])
    detail.append(["01/03/2025", "4000", "Sales invoice 101", "", "1000"])
    detail.append(["05/03/2025", "5000", "Purchase invoice 55", "500", ""])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_multisheet_workbook_expands_into_one_upload_per_sheet(http_client):
    """A VAT return export with separate Summary/Detail tabs used to be
    silently reduced to whichever sheet pandas reads by default - every
    sheet should now become its own classified sub-upload."""
    c = http_client
    practice_id = _signup(c, admin_email="multisheet@acme.test")
    resp = c.post(f"/practices/{practice_id}/clients", data={"name": "Multisheet Client"}, follow_redirects=False)
    client_id = resp.headers["location"].rsplit("/", 1)[-1]
    resp = c.post(f"/clients/{client_id}/jobs", data={
        "current_period_start": "2025-01-01", "current_period_end": "2025-12-31",
    }, follow_redirects=False)
    job_id = resp.headers["location"].rsplit("/", 1)[-1]

    resp = c.post(
        f"/jobs/{job_id}/uploads",
        files={"files": ("vat_return.xlsx", _make_multisheet_vat_workbook(),
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        follow_redirects=False,
    )
    assert resp.status_code == 303

    from app import storage
    job = storage.get_job(job_id)
    assert len(job["uploads"]) == 2
    by_sheet = {u["sheet_name"]: u for u in job["uploads"].values()}
    assert by_sheet["Summary"]["report_type"] == "vat_return"
    assert by_sheet["Detail"]["report_type"] == "nominal_activity"
    assert by_sheet["Summary"]["display_name"] == "vat_return.xlsx (Summary)"


def test_report_notes_and_section_scoped_upload(http_client):
    """The job page shows a fixed section per report type, each with an
    editable instruction note (persisted per client, not per job) and its
    own upload form - a file dropped into a section is classified as that
    section's report type directly, no guessing needed for that part."""
    c = http_client
    practice_id = _signup(c, admin_email="sections@acme.test")
    resp = c.post(f"/practices/{practice_id}/clients", data={"name": "Sections Client"}, follow_redirects=False)
    client_id = resp.headers["location"].rsplit("/", 1)[-1]
    resp = c.post(f"/clients/{client_id}/jobs", data={
        "current_period_start": "2025-01-01", "current_period_end": "2025-12-31",
    }, follow_redirects=False)
    job_id = resp.headers["location"].rsplit("/", 1)[-1]

    resp = c.get(f"/jobs/{job_id}")
    assert resp.status_code == 200
    assert "VAT Return" in resp.text
    assert "Add VAT Return file(s)" in resp.text

    resp = c.post(f"/clients/{client_id}/notes/vat_return", data={
        "note": "Detail tab has the transaction-level data.", "next": f"/jobs/{job_id}",
    }, follow_redirects=False)
    assert resp.status_code == 303
    assert resp.headers["location"] == f"/jobs/{job_id}"

    from app import storage
    client = storage.get_client(client_id)
    assert client["report_notes"]["vat_return"] == "Detail tab has the transaction-level data."
    assert "Instruction note (set)" in c.get(f"/jobs/{job_id}").text

    bank_path = SAMPLE_DIR / "bank_statement_current.csv"
    resp = c.post(
        f"/jobs/{job_id}/uploads",
        data={"section_report_type": "bank_statement"},
        files={"files": (bank_path.name, bank_path.read_bytes(), "text/csv")},
        follow_redirects=False,
    )
    assert resp.status_code == 303
    job = storage.get_job(job_id)
    upload = next(iter(job["uploads"].values()))
    assert upload["report_type"] == "bank_statement"


def test_ai_reconciliation_note_reaches_the_generated_workbook(http_client, monkeypatch):
    """Full wiring, end to end: template config opts in -> generate() calls
    the (mocked) agent for a flagged check -> the note lands on that
    check's sheet in the actual downloaded workbook. The Anthropic API is
    never really called in tests - see test_reconciliation_agent.py for
    the agent's own unit tests; this just proves main.py -> excel_builder
    wiring works, using a mock at the same patch point."""
    from unittest.mock import MagicMock, patch

    c = http_client
    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-key")
    practice_id = _signup(c, admin_email="ainotes@acme.test")

    resp = c.post(
        f"/practices/{practice_id}/templates",
        data={"name": "AI Notes Template"},
        files={"file": ("template.xlsx", _make_template_bytes(),
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        follow_redirects=False,
    )
    from app import storage
    template = storage.list_templates(practice_id)[0]
    template["config"]["ai_reconciliation_notes"]["enabled"] = True
    storage.save_template(template)

    resp = c.post(f"/practices/{practice_id}/clients", data={
        "name": "AI Notes Client", "template_id": template["id"],
    }, follow_redirects=False)
    client_id = resp.headers["location"].rsplit("/", 1)[-1]
    resp = c.post(f"/clients/{client_id}/jobs", data={
        "current_period_start": "2025-01-01", "current_period_end": "2025-12-31",
    }, follow_redirects=False)
    job_id = resp.headers["location"].rsplit("/", 1)[-1]

    resp = c.post(f"/clients/{client_id}/notes/vat_return", data={
        "note": "Detail tab has the transaction-level data.", "next": f"/jobs/{job_id}",
    }, follow_redirects=False)
    assert resp.status_code == 303

    # a VAT return with no TB/P&L uploaded alongside it - guaranteed to
    # flag (turnover and the VAT control balance both default to 0)
    vat_csv = b"Box 1,Box 2,Box 3,Box 4,Box 5,Box 6,Box 7,Box 8,Box 9\n1000,0,1000,200,800,15000,5000,0,0\n"
    resp = c.post(
        f"/jobs/{job_id}/uploads", data={"section_report_type": "vat_return"},
        files={"files": ("vat.csv", vat_csv, "text/csv")}, follow_redirects=False,
    )
    location = resp.headers["location"]
    assert "/mapping" in location
    upload_id = location.rsplit("/", 2)[-2]
    job = storage.get_job(job_id)
    upload = job["uploads"][upload_id]
    # real form submissions resubmit the pre-filled (auto-suggested)
    # column mapping; a bare {"action": "confirm"} would silently map
    # every column to None, zeroing out the VAT figures and defeating
    # the point of this test (nothing would ever flag).
    confirm_data = {"action": "confirm"}
    confirm_data.update({f"col__{col}": field for col, field in (upload["mapping"] or {}).items() if field})
    resp = c.post(location, data=confirm_data, follow_redirects=False)
    assert resp.status_code == 303

    mock_client = MagicMock()
    block = MagicMock()
    block.type = "text"
    block.text = "The £800 Box 5 figure has no nominal ledger postings to compare against - nothing was uploaded."
    response = MagicMock()
    response.content = [block]
    mock_client.messages.create.return_value = response

    with patch("anthropic.Anthropic", return_value=mock_client):
        resp = c.post(f"/jobs/{job_id}/generate", follow_redirects=False)
    assert resp.status_code == 303
    assert mock_client.messages.create.called

    resp = c.get(f"/jobs/{job_id}/download")
    assert resp.status_code == 200

    import io as _io

    import openpyxl
    wb = openpyxl.load_workbook(_io.BytesIO(resp.content))
    vat_sheet_names = [s for s in wb.sheetnames if "VAT" in s.upper()]
    assert vat_sheet_names, wb.sheetnames
    sheet_text = "\n".join(
        str(cell.value) for row in wb[vat_sheet_names[0]].iter_rows() for cell in row if cell.value
    )
    assert "AI-ASSISTED NOTE" in sheet_text
    assert "no nominal ledger postings" in sheet_text


def test_generate_progress_stream_reports_all_ten_steps_in_order(http_client):
    """The SSE progress endpoint (added so the person generating a
    working paper can see what's happening instead of staring at a
    blocked page) must report every one of the ten real steps, in the
    right order, ending with a completion event - and must leave the job
    in exactly the same generated state as the classic POST route does."""
    c = http_client
    practice_id = _signup(c, admin_email="progress@acme.test")
    resp = c.post(f"/practices/{practice_id}/clients", data={"name": "Progress Client"}, follow_redirects=False)
    client_id = resp.headers["location"].rsplit("/", 1)[-1]
    resp = c.post(f"/clients/{client_id}/jobs", data={
        "current_period_start": "2025-01-01", "current_period_end": "2025-12-31",
    }, follow_redirects=False)
    job_id = resp.headers["location"].rsplit("/", 1)[-1]

    resp = c.get(f"/jobs/{job_id}/generate/progress")
    assert resp.status_code == 200
    assert "EventSource" in resp.text

    resp = c.get(f"/jobs/{job_id}/generate/stream")
    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith("text/event-stream")

    import json as _json
    events = [
        _json.loads(line[len("data: "):])
        for line in resp.text.split("\n") if line.startswith("data: ")
    ]
    assert events[-1] == {"step": "complete", "redirect": f"/jobs/{job_id}"}

    from app.main import GENERATE_STEPS
    step_events = [e for e in events[:-1] if e["step"] != 5]  # step 5 (AI notes) is off by default -> "skipped" only
    seen_steps = [e["step"] for e in step_events]
    assert seen_steps == sorted(seen_steps)  # strictly non-decreasing: steps arrive in order
    for n in range(1, len(GENERATE_STEPS) + 1):
        if n == 5:
            assert any(e["step"] == 5 and e["status"] == "skipped" for e in events)
            continue
        assert any(e["step"] == n and e["status"] == "running" for e in events)
        assert any(e["step"] == n and e["status"] == "done" for e in events)

    from app import storage
    job = storage.get_job(job_id)
    assert job["status"] == "generated"
    assert job["output_file_id"]


def test_generate_stream_unauthenticated_redirects_to_login(http_client):
    c = http_client
    c.cookies.clear()
    resp = c.get("/jobs/job_doesnotmatter/generate/stream", follow_redirects=False)
    assert resp.status_code == 303
    assert resp.headers["location"].startswith("/login")


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
