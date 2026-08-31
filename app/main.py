import io
import json
import time
from datetime import date, datetime, timezone
from pathlib import Path

import pandas as pd
from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import RedirectResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app import accruals_prepayments, anomaly_detection, auth, brightpay_reports, compliance_checks, control_accounts, corporation_tax, document_detection, financial_statements, fixed_assets, going_concern, mapping, nominal_matrix, parsers, paye_reconciliation, recon, reconciliation_agent, related_party_transactions, statutory_deadlines, storage, vat_reconciliation, xero_reports
from app.excel_builder import build_workbook, build_workbook_into_template
from app.models import PAYE_RECON_TYPES, PERIODS, PLATFORMS, REPORT_LABELS, REPORT_SCHEMAS, REPORT_TYPES, REQUIRED_FIELDS, VAT_RECON_TYPES

BASE_DIR = Path(__file__).resolve().parent
app = FastAPI(title="Working Paper Automation")
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


@app.exception_handler(auth.Unauthenticated)
async def _unauthenticated_handler(request: Request, exc: auth.Unauthenticated):
    return RedirectResponse(f"/login?next={request.url.path}", status_code=303)


@app.exception_handler(auth.Forbidden)
async def _forbidden_handler(request: Request, exc: auth.Forbidden):
    user = auth.get_current_user(request)
    return templates.TemplateResponse(
        "error.html", {"request": request, "current_user": user, "message": exc.message}, status_code=403,
    )


def _authorize_practice(user: dict, practice_id: str) -> dict:
    practice = storage.get_practice(practice_id)
    if not practice or user["practice_id"] != practice_id:
        raise HTTPException(status_code=404)
    return practice


def _authorize_client(user: dict, client_id: str) -> dict:
    client = storage.get_client(client_id)
    if not client or client["practice_id"] != user["practice_id"]:
        raise HTTPException(status_code=404)
    if user["role"] == "preparer" and not storage.has_client_access(user["id"], client_id):
        raise auth.Forbidden("You don't have access to this client.")
    return client


def _authorize_job(user: dict, job_id: str) -> tuple[dict, dict]:
    job = storage.get_job(job_id)
    if not job:
        raise HTTPException(status_code=404)
    client = _authorize_client(user, job["client_id"])
    return job, client

# Canonical field(s) a given (report_type, period) upload feeds into the recon engine under.
DATA_KEY = {
    ("trial_balance", "current"): "tb_current",
    ("trial_balance", "comparative"): "tb_comparative",
    ("nominal_activity", "current"): "nominal_current",
    ("nominal_activity", "comparative"): "nominal_comparative",
    ("aged_debtors", "current"): "aged_debtors",
    ("aged_creditors", "current"): "aged_creditors",
    ("vat_return", "current"): "vat_return",
    ("bank_statement", "current"): "bank_statement",
    ("profit_and_loss", "current"): "pl_current",
    ("profit_and_loss", "comparative"): "pl_comparative",
    ("balance_sheet", "current"): "bs_current",
    ("balance_sheet", "comparative"): "bs_comparative",
    ("fixed_asset_register", "current"): "fixed_asset_register",
    ("fixed_asset_register", "comparative"): "fixed_asset_register",
    # vat_gl / vat_filed_sales / vat_filed_purchases are handled separately
    # in _load_canonical_data - every confirmed upload of one of those
    # types is concatenated (not just the latest one under a single key),
    # since the VAT Reconciliation workspace expects up to 10 filed-return
    # files combined into one dataset.
}

# section_report_type type_hints accepted on /jobs/{id}/uploads - the
# generic per-report-type sections (REPORT_TYPES) plus the VAT
# Reconciliation workspace's three dedicated upload zones. Kept as a
# separate list rather than folding VAT_RECON_TYPES into REPORT_TYPES so
# the general auto-detect classifier (document_detection.classify_report_type,
# only reached when no type_hint is given) never scores an unclassified
# bulk upload against these VAT-specific schemas.
UPLOADABLE_TYPES = REPORT_TYPES + VAT_RECON_TYPES


@app.get("/")
def home():
    return RedirectResponse("/practices")


@app.get("/practices")
def list_practices(request: Request):
    user = auth.get_current_user(request)
    if user:
        return RedirectResponse(f"/practices/{user['practice_id']}", status_code=303)
    # anonymous: not a directory of every practice (that would leak tenant
    # names to the world) - just the log-in / create-a-practice landing page.
    return templates.TemplateResponse("practices.html", {"request": request, "current_user": None})


@app.post("/practices")
def create_practice(
    request: Request,
    practice_name: str = Form(...), admin_name: str = Form(...),
    admin_email: str = Form(...), admin_password: str = Form(...),
):
    if storage.get_user_by_email(admin_email):
        return templates.TemplateResponse("practices.html", {
            "request": request, "current_user": None, "error": "That email is already registered - log in instead.",
        }, status_code=400)

    practice = storage.create_practice(practice_name.strip())
    password_hash = auth.hash_password(admin_password)
    user = storage.create_user(practice["id"], admin_email, password_hash, admin_name.strip(), "partner")

    response = RedirectResponse(f"/practices/{practice['id']}", status_code=303)
    auth.set_session_cookie(response, request, user["id"])
    return response


@app.get("/login")
def login_form(request: Request, next: str = "/practices"):
    user = auth.get_current_user(request)
    if user:
        return RedirectResponse(auth.safe_next_path(next), status_code=303)
    return templates.TemplateResponse("login.html", {"request": request, "current_user": None, "next": next})


@app.post("/login")
def login_submit(request: Request, email: str = Form(...), password: str = Form(...), next: str = Form("/practices")):
    user = storage.get_user_by_email(email)
    if not user or not auth.verify_password(password, user["password_hash"]):
        return templates.TemplateResponse("login.html", {
            "request": request, "current_user": None, "next": next, "error": "Invalid email or password.",
        }, status_code=400)
    response = RedirectResponse(auth.safe_next_path(next), status_code=303)
    auth.set_session_cookie(response, request, user["id"])
    return response


@app.post("/logout")
def logout():
    response = RedirectResponse("/login", status_code=303)
    response.delete_cookie(auth.SESSION_COOKIE)
    return response


@app.get("/practices/{practice_id}")
def practice_detail(request: Request, practice_id: str, user: dict = Depends(auth.current_user_dep)):
    practice = _authorize_practice(user, practice_id)
    templates_list = storage.list_templates(practice_id)
    return templates.TemplateResponse("practice_detail.html", {
        "request": request, "current_user": user, "practice": practice, "templates_list": templates_list,
        "breadcrumbs": [{"label": practice["name"]}],
    })


@app.get("/practices/{practice_id}/clients")
def practice_clients(request: Request, practice_id: str, user: dict = Depends(auth.current_user_dep)):
    practice = _authorize_practice(user, practice_id)
    templates_list = storage.list_templates(practice_id)
    clients = storage.list_clients(practice_id)
    if user["role"] == "preparer":
        accessible = set(storage.list_client_access(user["id"]))
        clients = [c for c in clients if c["id"] in accessible]
    template_names = {t["id"]: t["name"] for t in templates_list}
    return templates.TemplateResponse("practice_clients.html", {
        "request": request, "current_user": user, "practice": practice, "templates_list": templates_list,
        "clients": clients, "template_names": template_names,
        "breadcrumbs": [{"label": practice["name"], "url": f"/practices/{practice_id}"}, {"label": "Clients"}],
    })


@app.post("/practices/{practice_id}/templates")
async def upload_template(practice_id: str, name: str = Form(...), file: UploadFile = None,
                           user: dict = Depends(auth.current_user_dep)):
    _authorize_practice(user, practice_id)
    auth.require_role(user, "partner", "manager")
    content = await file.read()
    storage.create_template(practice_id, name.strip(), content, file.filename)
    return RedirectResponse(f"/practices/{practice_id}", status_code=303)


@app.get("/practices/{practice_id}/templates/{template_id}")
def template_detail(request: Request, practice_id: str, template_id: str, user: dict = Depends(auth.current_user_dep)):
    import json as _json
    practice = _authorize_practice(user, practice_id)
    auth.require_role(user, "partner", "manager")
    template = storage.get_template(practice_id, template_id)
    return templates.TemplateResponse("template_detail.html", {
        "request": request, "current_user": user, "practice_id": practice_id, "template": template,
        "config_json": _json.dumps(template["config"], indent=2),
        "breadcrumbs": [{"label": practice["name"], "url": f"/practices/{practice_id}"}, {"label": template["name"]}],
    })


@app.post("/practices/{practice_id}/templates/{template_id}/config")
async def save_template_config(request: Request, practice_id: str, template_id: str,
                                user: dict = Depends(auth.current_user_dep)):
    import json as _json
    _authorize_practice(user, practice_id)
    auth.require_role(user, "partner", "manager")
    form = await request.form()
    template = storage.get_template(practice_id, template_id)
    try:
        template["config"] = _json.loads(form.get("config_json"))
    except _json.JSONDecodeError as exc:
        return templates.TemplateResponse("template_detail.html", {
            "request": request, "current_user": user, "practice_id": practice_id, "template": template,
            "config_json": form.get("config_json"), "error": f"Invalid JSON: {exc}",
        })
    storage.save_template(template)
    return RedirectResponse(f"/practices/{practice_id}/templates/{template_id}", status_code=303)


@app.post("/practices/{practice_id}/templates/{template_id}/make-default")
def make_template_default(practice_id: str, template_id: str, user: dict = Depends(auth.current_user_dep)):
    practice = _authorize_practice(user, practice_id)
    auth.require_role(user, "partner", "manager")
    practice["default_template_id"] = template_id
    storage.save_practice(practice)
    return RedirectResponse(f"/practices/{practice_id}", status_code=303)


@app.post("/practices/{practice_id}/clients")
def create_client(practice_id: str, name: str = Form(...), template_id: str = Form(""),
                   user: dict = Depends(auth.current_user_dep)):
    practice = _authorize_practice(user, practice_id)
    auth.require_role(user, "partner", "manager")
    chosen_template = template_id or practice.get("default_template_id")
    client = storage.create_client(practice_id, name.strip(), chosen_template)
    return RedirectResponse(f"/clients/{client['id']}", status_code=303)


@app.get("/practices/{practice_id}/users")
def list_users(request: Request, practice_id: str, user: dict = Depends(auth.current_user_dep)):
    practice = _authorize_practice(user, practice_id)
    auth.require_role(user, "partner")
    practice_users = storage.list_users(practice_id)
    clients = storage.list_clients(practice_id)
    access_by_user = {u["id"]: set(storage.list_client_access(u["id"])) for u in practice_users if u["role"] == "preparer"}
    return templates.TemplateResponse("users.html", {
        "request": request, "current_user": user, "practice_id": practice_id,
        "practice_users": practice_users, "clients": clients, "access_by_user": access_by_user,
        "roles": auth.ROLES,
        "breadcrumbs": [{"label": practice["name"], "url": f"/practices/{practice_id}"}, {"label": "Users"}],
    })


@app.post("/practices/{practice_id}/users")
def create_user(practice_id: str, name: str = Form(...), email: str = Form(...),
                 password: str = Form(...), role: str = Form(...),
                 user: dict = Depends(auth.current_user_dep)):
    _authorize_practice(user, practice_id)
    auth.require_role(user, "partner")
    if role not in auth.ROLES:
        raise HTTPException(status_code=400, detail="Invalid role")
    if storage.get_user_by_email(email):
        raise auth.Forbidden("That email is already registered.")
    password_hash = auth.hash_password(password)
    storage.create_user(practice_id, email, password_hash, name.strip(), role)
    return RedirectResponse(f"/practices/{practice_id}/users", status_code=303)


@app.post("/practices/{practice_id}/users/{user_id}/client-access")
async def set_user_client_access(request: Request, practice_id: str, user_id: str,
                                  user: dict = Depends(auth.current_user_dep)):
    _authorize_practice(user, practice_id)
    auth.require_role(user, "partner")
    target = storage.get_user(user_id)
    if not target or target["practice_id"] != practice_id:
        raise HTTPException(status_code=404)
    form = await request.form()
    client_ids = form.getlist("client_ids")
    storage.set_client_access(user_id, client_ids)
    return RedirectResponse(f"/practices/{practice_id}/users", status_code=303)


@app.post("/practices/{practice_id}/users/{user_id}/delete")
def delete_user(practice_id: str, user_id: str, user: dict = Depends(auth.current_user_dep)):
    _authorize_practice(user, practice_id)
    auth.require_role(user, "partner")
    target = storage.get_user(user_id)
    if not target or target["practice_id"] != practice_id:
        raise HTTPException(status_code=404)
    if target["id"] == user["id"]:
        raise auth.Forbidden("You can't remove your own account.")
    storage.delete_user(user_id)
    return RedirectResponse(f"/practices/{practice_id}/users", status_code=303)


@app.get("/clients/{client_id}")
def client_detail(request: Request, client_id: str, user: dict = Depends(auth.current_user_dep)):
    client = _authorize_client(user, client_id)
    jobs = storage.list_jobs(client_id)
    practice = storage.get_practice(client["practice_id"])
    return templates.TemplateResponse("client_detail.html", {
        "request": request, "current_user": user, "client": client, "jobs": jobs,
        "breadcrumbs": [
            {"label": practice["name"], "url": f"/practices/{client['practice_id']}"},
            {"label": "Clients", "url": f"/practices/{client['practice_id']}/clients"},
            {"label": client["name"]},
        ],
    })


@app.post("/clients/{client_id}/jobs")
def create_job(
    client_id: str,
    current_period_start: str = Form(...), current_period_end: str = Form(...),
    comparative_period_start: str = Form(""), comparative_period_end: str = Form(""),
    user: dict = Depends(auth.current_user_dep),
):
    client = _authorize_client(user, client_id)
    job = storage.create_job(
        client_id, client["name"],
        current_period_start, current_period_end,
        comparative_period_start or None, comparative_period_end or None,
    )
    return RedirectResponse(f"/jobs/{job['id']}", status_code=303)


@app.get("/jobs/{job_id}")
def job_detail(request: Request, job_id: str, user: dict = Depends(auth.current_user_dep)):
    job, client = _authorize_job(user, job_id)
    practice = storage.get_practice(client["practice_id"])
    uploads_by_type = {rt: [] for rt in REPORT_TYPES}
    uploads_by_type.setdefault("", [])  # couldn't be classified at all
    for upload in job["uploads"].values():
        uploads_by_type.setdefault(upload["report_type"], []).append(upload)
    return templates.TemplateResponse("job_detail.html", {
        "request": request, "current_user": user, "job": job, "client": client,
        "report_types": REPORT_TYPES, "report_labels": REPORT_LABELS,
        "platforms": PLATFORMS, "periods": PERIODS,
        "uploads_by_type": uploads_by_type,
        "report_notes": client.get("report_notes", {}),
        "progress_summary": _summarize_progress(job.get("progress")),
        "vat_recon_types": VAT_RECON_TYPES,
        "paye_recon_types": PAYE_RECON_TYPES,
        "breadcrumbs": [
            {"label": practice["name"], "url": f"/practices/{client['practice_id']}"},
            {"label": "Clients", "url": f"/practices/{client['practice_id']}/clients"},
            {"label": client["name"], "url": f"/clients/{client['id']}"},
            {"label": job["current_label"]},
        ],
    })


@app.post("/jobs/{job_id}/tax-inputs")
def save_tax_inputs(
    job_id: str,
    ct_associated_companies: int = Form(0),
    ct_disallowable_additions: float = Form(0.0),
    ct_capital_allowances: float = Form(0.0),
    user: dict = Depends(auth.current_user_dep),
):
    job, _client = _authorize_job(user, job_id)
    job["ct_associated_companies"] = ct_associated_companies
    job["ct_disallowable_additions"] = ct_disallowable_additions
    job["ct_capital_allowances"] = ct_capital_allowances
    storage.save_job(job)
    return RedirectResponse(f"/jobs/{job_id}", status_code=303)


@app.post("/jobs/{job_id}/vat-recon-settings")
def save_vat_recon_settings(
    job_id: str,
    vat_recon_basis: str = Form("accrual"),
    vat_recon_tolerance: float = Form(0.0),
    user: dict = Depends(auth.current_user_dep),
):
    job, _client = _authorize_job(user, job_id)
    job["vat_recon_basis"] = vat_recon_basis if vat_recon_basis in ("accrual", "cash") else "accrual"
    job["vat_recon_tolerance"] = max(0.0, vat_recon_tolerance)
    storage.save_job(job)
    return RedirectResponse(f"/jobs/{job_id}#vat-recon", status_code=303)


def _jsonable(value):
    """A DataFrame cell can hold a pandas Timestamp, a numpy scalar, NaN/NaT,
    or a plain Python value - job["..."] is stored as Postgres JSONB (see
    storage._put_entity), which only accepts the plain kind, same reason
    job["summary"]/job["progress"] are built from plain dicts rather than
    passing DataFrame rows straight through."""
    if value is None:
        return None
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d") if not pd.isna(value) else None
    if hasattr(value, "item"):  # numpy scalar (float64/int64/bool_)
        value = value.item()
    if isinstance(value, float) and pd.isna(value):
        return None
    return value


def _df_to_records(df: pd.DataFrame) -> list[dict]:
    if df is None or df.empty:
        return []
    return [{col: _jsonable(v) for col, v in row.items()} for row in df.to_dict(orient="records")]


def _recon_result_to_dict(result) -> dict:
    return {
        "name": result.name, "status": result.status, "message": result.message,
        "detail": _df_to_records(result.detail),
        "extra_detail": _df_to_records(result.extra_detail),
        "extra_detail_label": result.extra_detail_label,
        "matched_detail": _df_to_records(getattr(result, "matched_detail", None)),
        "matched_detail_label": getattr(result, "matched_detail_label", ""),
    }


def _control_account_result_to_dict(result) -> dict:
    """Same {name/status/message/detail/extra_detail/matched_detail} shape
    _recon_result_to_dict produces, so job_detail.html's one results-
    rendering macro handles both without knowing the difference - even
    though ControlAccountResult is a different dataclass with its own
    field names (schedule/breakdown/extra_detail, not detail/extra_detail/
    matched_detail). breakdown (the aged-listing tie-out) is preferred for
    the "extra_detail" slot when present; movement_breakdown (the net-
    movement-by-contact view every OTHER control account gets instead)
    fills the same slot when there's no aged listing to show - a control
    account always has one or the other, never both, so nothing is lost
    picking whichever is populated."""
    breakdown, breakdown_label = result.breakdown, result.breakdown_label
    if breakdown is None or breakdown.empty:
        breakdown, breakdown_label = result.movement_breakdown, result.movement_breakdown_label
    return {
        "name": f"{result.account_code} – {result.account_name}", "status": result.status, "message": result.message,
        "detail": _df_to_records(result.schedule),
        "extra_detail": _df_to_records(breakdown),
        "extra_detail_label": breakdown_label,
        "matched_detail": _df_to_records(result.extra_detail),
        "matched_detail_label": result.extra_detail_label,
    }


def _asset_register_result_to_dict(result) -> dict:
    """AssetRegisterResult (fixed_assets.asset_level_rollforward) has no
    `name` field, unlike FixedAssetResult/ReconResult (it's always exactly
    one register per job, so nothing to name it against), and four data
    tables instead of the generic detail/extra_detail/matched_detail three
    slots this shape carries. Mapped here as: detail = the rolled-forward
    per-asset schedule (the main content), extra_detail = new additions
    found in nominal activity not yet in the register, matched_detail =
    possible disposals. `summary`'s few totals (register NBV vs TB,
    variance) are already stated in `message`, so it isn't duplicated as
    its own table."""
    return {
        "name": "Fixed asset register (asset detail)", "status": result.status, "message": result.message,
        "detail": _df_to_records(result.asset_schedule),
        "extra_detail": _df_to_records(result.new_additions),
        "extra_detail_label": "New additions found in nominal activity, not yet in the register",
        "matched_detail": _df_to_records(result.possible_disposals),
        "matched_detail_label": "Possible disposals (credit movements on fixed asset cost codes)",
    }


@app.post("/jobs/{job_id}/vat-recon/run")
def run_vat_reconciliation(job_id: str, user: dict = Depends(auth.current_user_dep)):
    """Computes the VAT Reconciliation independently of the main Generate
    pipeline - this is its own section, testable and reviewable on its
    own, before a full working paper is even ready to build (see
    app/vat_reconciliation.py for the matching logic itself)."""
    job, _client = _authorize_job(user, job_id)
    data = _load_canonical_data(job)
    settings = vat_reconciliation.VatReconSettings(
        accounting_basis=job.get("vat_recon_basis", "accrual"),
        tolerance=job.get("vat_recon_tolerance", 0.0),
    )
    results = vat_reconciliation.reconcile(data, settings)
    job["vat_recon_results"] = [_recon_result_to_dict(r) for r in results]
    job["vat_recon_computed_at"] = datetime.now(timezone.utc).isoformat()
    storage.save_job(job)
    return RedirectResponse(f"/jobs/{job_id}#vat-recon", status_code=303)


@app.post("/jobs/{job_id}/paye-recon-settings")
def save_paye_recon_settings(
    job_id: str,
    paye_recon_tolerance: float = Form(0.0),
    paye_recon_date_window_days: int = Form(paye_reconciliation.DEFAULT_DATE_WINDOW_DAYS),
    user: dict = Depends(auth.current_user_dep),
):
    job, _client = _authorize_job(user, job_id)
    job["paye_recon_tolerance"] = max(0.0, paye_recon_tolerance)
    job["paye_recon_date_window_days"] = max(0, paye_recon_date_window_days)
    storage.save_job(job)
    return RedirectResponse(f"/jobs/{job_id}#paye-recon", status_code=303)


@app.post("/jobs/{job_id}/paye-recon/run")
def run_paye_reconciliation(job_id: str, user: dict = Depends(auth.current_user_dep)):
    """Computes the PAYE Reconciliation independently of the main Generate
    pipeline, same reasoning as VAT Reconciliation above - reuses the
    job's existing General Ledger upload (nominal_current) rather than
    needing one of its own (see app/paye_reconciliation.py)."""
    job, _client = _authorize_job(user, job_id)
    data = _load_canonical_data(job)
    settings = paye_reconciliation.PayeReconSettings(
        tolerance=job.get("paye_recon_tolerance", 0.0),
        date_window_days=job.get("paye_recon_date_window_days", paye_reconciliation.DEFAULT_DATE_WINDOW_DAYS),
    )
    results = paye_reconciliation.reconcile(data, settings)
    job["paye_recon_results"] = [_recon_result_to_dict(r) for r in results]
    job["paye_recon_computed_at"] = datetime.now(timezone.utc).isoformat()
    storage.save_job(job)
    return RedirectResponse(f"/jobs/{job_id}#paye-recon", status_code=303)


@app.post("/jobs/{job_id}/control-accounts/run")
def run_control_accounts(job_id: str, user: dict = Depends(auth.current_user_dep)):
    """Computes Control Account rollforwards independently of the main
    Generate pipeline, same "test this section on its own" treatment as
    VAT/PAYE Reconciliation above - reuses whatever Trial Balance/Nominal
    Activity/Aged Debtors/Aged Creditors uploads are already confirmed
    for this job rather than needing uploads of its own (see
    app/control_accounts.py)."""
    job, client = _authorize_job(user, job_id)
    data = _load_canonical_data(job)
    template = storage.get_template(client["practice_id"], client["template_id"]) if client.get("template_id") else None
    materiality, _ = _job_materiality(template)
    results = control_accounts.build_all_rollforwards(
        data.get("tb_current"), data.get("tb_comparative"), data.get("nominal_current"),
        data.get("aged_debtors"), data.get("aged_creditors"), materiality,
    )
    miscoding = control_accounts.suggest_control_account_miscoding(
        data.get("tb_current"), data.get("nominal_current"),
        [(r.account_code, r.account_name) for r in results],
        data.get("aged_debtors"), data.get("aged_creditors"), materiality,
    )
    job["control_accounts_results"] = [_control_account_result_to_dict(r) for r in results] + [_recon_result_to_dict(miscoding)]
    job["control_accounts_computed_at"] = datetime.now(timezone.utc).isoformat()
    storage.save_job(job)
    return RedirectResponse(f"/jobs/{job_id}#control-accounts", status_code=303)


@app.post("/jobs/{job_id}/debtors-creditors/run")
def run_debtors_creditors_recon(job_id: str, user: dict = Depends(auth.current_user_dep)):
    """Computes the aged debtors/creditors listing vs Trial Balance
    control-account check independently of the main Generate pipeline,
    same "test this section on its own" treatment as VAT/PAYE
    Reconciliation and Control Accounts above - reuses whatever Aged
    Debtors/Aged Creditors/Trial Balance uploads are already confirmed
    for this job (see recon.debtors_creditors_control_recon)."""
    job, client = _authorize_job(user, job_id)
    data = _load_canonical_data(job)
    template = storage.get_template(client["practice_id"], client["template_id"]) if client.get("template_id") else None
    materiality, _ = _job_materiality(template)
    results = [
        recon.debtors_creditors_control_recon(
            data.get("aged_debtors"), data.get("tb_current"),
            ["debtors control", "trade debtors", "accounts receivable", "sales ledger control"],
            "customer", "Debtors control account reconciliation", materiality,
        ),
        recon.debtors_creditors_control_recon(
            data.get("aged_creditors"), data.get("tb_current"),
            ["creditors control", "trade creditors", "accounts payable", "purchase ledger control"],
            "supplier", "Creditors control account reconciliation", materiality,
        ),
    ]
    job["debtors_creditors_results"] = [_recon_result_to_dict(r) for r in results]
    job["debtors_creditors_computed_at"] = datetime.now(timezone.utc).isoformat()
    storage.save_job(job)
    return RedirectResponse(f"/jobs/{job_id}#debtors-creditors", status_code=303)


@app.post("/jobs/{job_id}/fixed-asset-register/run")
def run_fixed_asset_register(job_id: str, user: dict = Depends(auth.current_user_dep)):
    """Computes the Fixed Asset Register checks independently of the main
    Generate pipeline, same "test this section on its own" treatment as
    VAT/PAYE Reconciliation, Control Accounts and Debtors & Creditors
    above (see app/fixed_assets.py). The category-level rollforward and
    the capex-miscoding suggestion need only the Trial Balance/Nominal
    Activity uploads already used elsewhere; the asset-level rollforward
    additionally needs a prior-year Fixed Asset Register upload - when
    that isn't present yet it comes back "n/a" rather than blocking the
    other two checks from running."""
    job, client = _authorize_job(user, job_id)
    data = _load_canonical_data(job)
    template = storage.get_template(client["practice_id"], client["template_id"]) if client.get("template_id") else None
    materiality, _ = _job_materiality(template)
    category_result = fixed_assets.category_level_rollforward(
        data.get("tb_current"), data.get("tb_comparative"), data.get("nominal_current"),
        period_days=_period_days(job), materiality=materiality, period_start=_period_start(job),
    )
    register_result = fixed_assets.asset_level_rollforward(
        data.get("fixed_asset_register"), data.get("nominal_current"), data.get("tb_current"),
        period_days=_period_days(job), materiality=materiality, period_start=_period_start(job),
    )
    capex_result = fixed_assets.suggest_capital_expenditure_reclassification(
        data.get("tb_current"), data.get("nominal_current"), data.get("fixed_asset_register"), threshold=materiality,
    )
    job["fixed_asset_register_results"] = [
        _recon_result_to_dict(category_result),
        _asset_register_result_to_dict(register_result),
        _recon_result_to_dict(capex_result),
    ]
    job["fixed_asset_register_computed_at"] = datetime.now(timezone.utc).isoformat()
    storage.save_job(job)
    return RedirectResponse(f"/jobs/{job_id}#fixed-asset-register", status_code=303)


@app.post("/jobs/{job_id}/bank-recon/run")
def run_bank_reconciliation(job_id: str, user: dict = Depends(auth.current_user_dep)):
    """Computes Bank Reconciliation independently of the main Generate
    pipeline, same "test this section on its own" treatment as every
    other standalone section above (see recon.bank_reconciliation) -
    reuses whatever Bank Closing Statement/Trial Balance uploads are
    already confirmed for this job rather than needing uploads of its
    own. Last of the four sections built for this phase (VAT/PAYE,
    Control Accounts/Debtors & Creditors, Fixed Asset Register, now
    Bank), so every check on the job page can be run and reviewed on
    its own before a full working paper is even ready to build."""
    job, client = _authorize_job(user, job_id)
    data = _load_canonical_data(job)
    template = storage.get_template(client["practice_id"], client["template_id"]) if client.get("template_id") else None
    materiality, _ = _job_materiality(template)
    result = recon.bank_reconciliation(data.get("bank_statement"), data.get("tb_current"), materiality=materiality)
    job["bank_recon_results"] = [_recon_result_to_dict(result)]
    job["bank_recon_computed_at"] = datetime.now(timezone.utc).isoformat()
    storage.save_job(job)
    return RedirectResponse(f"/jobs/{job_id}#bank-recon", status_code=303)


@app.post("/clients/{client_id}/notes/{report_type}")
def save_report_note(client_id: str, report_type: str, note: str = Form(""), next: str = Form(""),
                      user: dict = Depends(auth.current_user_dep)):
    client = _authorize_client(user, client_id)
    if report_type not in REPORT_TYPES:
        raise HTTPException(status_code=400)
    notes = dict(client.get("report_notes") or {})
    if note.strip():
        notes[report_type] = note.strip()
    else:
        notes.pop(report_type, None)
    client["report_notes"] = notes
    storage.save_client(client)
    destination = next if next.startswith("/") and not next.startswith("//") else f"/clients/{client_id}"
    return RedirectResponse(destination, status_code=303)


def _add_ai_reconciliation_notes(results: list, client: dict, job: dict) -> None:
    """Mutates each flagged (review/error) result in place, adding a short
    AI-assisted note - see app/reconciliation_agent.py for what it's given
    and the guardrails around it. Called only when the practice has opted
    in for this template (see generate()); any failure here degrades to no
    note per-check, never breaks generation."""
    notes = client.get("report_notes") or {}
    notes_context = "\n".join(f"- {REPORT_LABELS.get(rt, rt)}: {note}" for rt, note in notes.items())
    period_label = job.get("current_label", "")
    for result in results:
        if result.status not in ("review", "error"):
            continue
        try:
            result.ai_note = reconciliation_agent.explain_flagged_result(
                result.name, result.status, result.message,
                result.detail, result.extra_detail, result.extra_detail_label,
                notes_context, job.get("client_name", ""), period_label,
            )
        except Exception:
            result.ai_note = ""


def _summarize_progress(progress: list[dict] | None) -> dict | None:
    """Turns the flat list of persisted generate events (see
    _generate_workbook_steps) into a per-step timing breakdown - a
    step's duration is just the gap between its own "running" and
    "done"/"skipped" timestamps, already there for the reading once
    events are persisted with a timestamp on each one. "retrying"
    events (from _run_step_with_retry, steps 1 and 10 only) don't move
    that gap - they're counted per step instead, so a step that needed
    a couple of attempts still shows its real end-to-end duration plus
    how many times it had to retry."""
    if not progress:
        return None

    by_step: dict[int, dict] = {}
    for evt in progress:
        step = evt.get("step")
        if not isinstance(step, int):
            continue
        entry = by_step.setdefault(step, {"step": step, "label": evt.get("label", ""), "retries": 0})
        if evt["status"] == "retrying":
            entry["retries"] += 1
            entry["last_retry_error"] = evt.get("error")
            continue
        entry[evt["status"]] = evt["at"]
        if "detail" in evt:
            entry["detail"] = evt["detail"]

    rows = []
    for step in sorted(by_step):
        entry = by_step[step]
        running_at, done_at, skipped_at = entry.get("running"), entry.get("done"), entry.get("skipped")
        seconds = None
        if running_at and done_at:
            seconds = (datetime.fromisoformat(done_at) - datetime.fromisoformat(running_at)).total_seconds()
        status = "done" if done_at else ("skipped" if skipped_at else "running")
        rows.append({
            "step": step, "label": entry["label"], "status": status, "seconds": seconds,
            "detail": entry.get("detail"), "retries": entry["retries"],
            "last_retry_error": entry.get("last_retry_error"),
        })

    error_evt = next((e for e in progress if e.get("step") == "error"), None)
    complete_evt = next((e for e in progress if e.get("step") == "complete"), None)
    total_seconds = None
    if complete_evt:
        total_seconds = (datetime.fromisoformat(complete_evt["at"]) - datetime.fromisoformat(progress[0]["at"])).total_seconds()

    return {
        "rows": rows, "total_seconds": total_seconds,
        "finished": complete_evt is not None, "error_message": error_evt["message"] if error_evt else None,
    }


def _first_unconfirmed_upload(job: dict) -> str | None:
    for upload_id, upload in job["uploads"].items():
        if not upload["confirmed"]:
            return upload_id
    return None


async def _ingest_one_upload(job_id: str, job: dict, filename: str, content: bytes,
                              sheet_name: str | None = None, display_name: str | None = None,
                              type_hint: str | None = None) -> None:
    """Classifies and stores a single uploaded file (or, for a multi-sheet
    workbook, a single sheet of it - see upload_files() below, which is
    what passes sheet_name/display_name). A genuine Xero-native match (the
    file's actual layout parses cleanly, not just similar-looking column
    names) auto-confirms immediately, same as before this feature existed
    - that's a structural validation, not a guess, so it wins even over a
    type_hint (uploading into the wrong section shouldn't force a
    misdetection). Anything else (including PDFs, which flow through the
    same DataSource - see parsers.py) gets queued for a quick human confirm
    rather than trusted blind: report type comes from type_hint when the
    file was dropped into a specific document-type section (see
    job_detail.html) rather than guessed, but platform/period are always
    guessed since the section doesn't tell us those."""
    source = parsers.FileDataSource(content, filename=filename, sheet_name=sheet_name)
    display_name = display_name or filename

    xero_report_type = document_detection.try_xero_native(source)
    if xero_report_type:
        period = document_detection.guess_period(source, xero_report_type, job, xero_report_type)
        expected_end = _expected_period_end(job, period)
        period_check = xero_reports.check_period(xero_reports.extract_period_info(source), expected_end)
        columns = source.raw_columns()
        upload_id = storage.add_upload(job, xero_report_type, period, "xero", filename, content, columns)
        job["uploads"][upload_id]["mapping"] = {}
        job["uploads"][upload_id]["confirmed"] = True
        job["uploads"][upload_id]["xero_native"] = True
        job["uploads"][upload_id]["period_check"] = period_check
        job["uploads"][upload_id]["sheet_name"] = sheet_name
        job["uploads"][upload_id]["display_name"] = display_name
        storage.save_job(job)
        return

    brightpay_report_type = brightpay_reports.try_brightpay_native(source)
    if brightpay_report_type:
        # Same reasoning as the Xero-native branch above: a genuine
        # structural match on BrightPay's fixed export shape is certain
        # enough to skip the confirm-mapping step entirely. BrightPay
        # reports don't have a single current/comparative period the way
        # a Xero TB does (one file spans many tax months), so "period" is
        # just a placeholder here - _load_canonical_data concatenates
        # every confirmed upload of these types together regardless of it.
        columns = source.raw_columns()
        upload_id = storage.add_upload(job, brightpay_report_type, "current", "brightpay", filename, content, columns)
        job["uploads"][upload_id]["mapping"] = {}
        job["uploads"][upload_id]["confirmed"] = True
        job["uploads"][upload_id]["sheet_name"] = sheet_name
        job["uploads"][upload_id]["display_name"] = display_name
        storage.save_job(job)
        return

    columns = source.raw_columns()
    if type_hint:
        report_type, confidence = type_hint, 1.0
    else:
        report_type, confidence = document_detection.classify_report_type(columns)
        if report_type in ("profit_and_loss", "balance_sheet"):
            profile_probe = mapping.suggest_mapping(report_type, columns)
            category_col = next((c for c, f in profile_probe.items() if f == "category"), None)
            report_type = document_detection.disambiguate_pl_vs_bs(source.raw_dataframe(), category_col)
    platform = document_detection.classify_platform(columns, is_xero_native=False)
    period = document_detection.guess_period(source, report_type or "trial_balance", job, None)

    upload_id = storage.add_upload(job, report_type or "", period, platform, filename, content, columns)

    suggestion = {}
    validation_note = ""
    if report_type:
        profile = mapping.load_profile(job["client_id"], report_type, platform)
        suggestion = {col: profile.get(col) for col in columns} if profile else mapping.suggest_mapping(report_type, columns)
        required = REQUIRED_FIELDS.get(report_type, [])
        matched_required = {v for v in suggestion.values() if v} & set(required)
        if required and not matched_required:
            validation_note = (
                f"None of this file's columns look like a {REPORT_LABELS[report_type]} - "
                f"double check the report type below before mapping columns."
            )
    else:
        validation_note = (
            "Couldn't confidently guess what kind of report this is - pick the report type below, "
            "then map its columns."
        )

    job["uploads"][upload_id]["mapping"] = suggestion
    job["uploads"][upload_id]["validation_note"] = validation_note
    job["uploads"][upload_id]["detection_confidence"] = confidence
    job["uploads"][upload_id]["sheet_name"] = sheet_name
    job["uploads"][upload_id]["display_name"] = display_name
    storage.save_job(job)


@app.post("/jobs/{job_id}/uploads")
async def upload_files(job_id: str, files: list[UploadFile] = File(...),
                        section_report_type: str = Form(""),
                        user: dict = Depends(auth.current_user_dep)):
    job, _client = _authorize_job(user, job_id)
    type_hint = section_report_type if section_report_type in UPLOADABLE_TYPES else None

    for file in files:
        content = await file.read()
        if not content:
            continue

        sheet_names = [None]
        if Path(file.filename).suffix.lower() in (".xlsx", ".xls"):
            try:
                found = parsers.excel_sheet_names(content)
            except Exception:
                found = []
            if len(found) > 1:
                # a workbook with several tabs (e.g. a VAT return export
                # with separate Summary and Detail sheets) would otherwise
                # be silently reduced to whichever sheet pandas reads by
                # default - every sheet becomes its own classified
                # sub-upload instead, through the exact same path as any
                # other file in this batch.
                sheet_names = found

        for sheet in sheet_names:
            job = storage.get_job(job_id)  # re-fetch: each ingest may have just saved the job
            display_name = f"{file.filename} ({sheet})" if sheet else file.filename
            await _ingest_one_upload(job_id, job, file.filename, content, sheet_name=sheet,
                                      display_name=display_name, type_hint=type_hint)

    job = storage.get_job(job_id)
    next_upload_id = _first_unconfirmed_upload(job)
    if next_upload_id:
        return RedirectResponse(f"/jobs/{job_id}/uploads/{next_upload_id}/mapping", status_code=303)
    return RedirectResponse(f"/jobs/{job_id}", status_code=303)


def _expected_period_end(job: dict, period: str):
    from datetime import date
    key = "current_period_end" if period == "current" else "comparative_period_end"
    value = job.get(key)
    return date.fromisoformat(value) if value else None


@app.get("/jobs/{job_id}/uploads/{upload_id}/mapping")
def mapping_form(request: Request, job_id: str, upload_id: str, user: dict = Depends(auth.current_user_dep)):
    job, _client = _authorize_job(user, job_id)
    upload = job["uploads"][upload_id]
    canonical_fields = REPORT_SCHEMAS.get(upload["report_type"], {})
    return templates.TemplateResponse("mapping.html", {
        "request": request, "current_user": user, "job": job, "upload": upload,
        "canonical_fields": canonical_fields,
        "report_label": REPORT_LABELS.get(upload["report_type"], "(report type not yet set)"),
        "report_types": REPORT_TYPES, "report_labels": REPORT_LABELS, "platforms": PLATFORMS, "periods": PERIODS,
    })


@app.post("/jobs/{job_id}/uploads/{upload_id}/mapping")
async def save_mapping(request: Request, job_id: str, upload_id: str, user: dict = Depends(auth.current_user_dep)):
    job, _client = _authorize_job(user, job_id)
    form = await request.form()
    upload = job["uploads"][upload_id]

    upload["report_type"] = form.get("report_type") or upload["report_type"]
    upload["platform"] = form.get("platform") or upload["platform"]
    upload["period"] = form.get("period") or upload["period"]

    if form.get("action") == "retype":
        # the report type (or platform) changed - re-suggest the column
        # mapping against the new schema rather than confirming stale
        # selections that were guessed for a different report type.
        profile = mapping.load_profile(job["client_id"], upload["report_type"], upload["platform"])
        upload["mapping"] = (
            {col: profile.get(col) for col in upload["columns"]} if profile
            else mapping.suggest_mapping(upload["report_type"], upload["columns"])
        )
        upload["validation_note"] = ""
        storage.save_job(job)
        return RedirectResponse(f"/jobs/{job_id}/uploads/{upload_id}/mapping", status_code=303)

    if not upload["report_type"]:
        upload["validation_note"] = "Pick a report type before confirming - without one, this file's data can't be mapped to anything."
        storage.save_job(job)
        return RedirectResponse(f"/jobs/{job_id}/uploads/{upload_id}/mapping", status_code=303)

    new_mapping = {}
    for col in upload["columns"]:
        key = f"col__{col}"
        val = form.get(key)
        new_mapping[col] = val if val else None

    upload["mapping"] = new_mapping
    upload["confirmed"] = True
    storage.save_job(job)

    if form.get("save_profile"):
        mapping.save_profile(job["client_id"], upload["report_type"], upload["platform"], new_mapping)

    next_upload_id = _first_unconfirmed_upload(job)
    if next_upload_id and next_upload_id != upload_id:
        return RedirectResponse(f"/jobs/{job_id}/uploads/{next_upload_id}/mapping", status_code=303)

    return RedirectResponse(f"/jobs/{job_id}", status_code=303)


def _load_canonical_data(job: dict) -> dict:
    data = {}
    for upload in job["uploads"].values():
        if not upload["confirmed"]:
            continue
        content = storage.load_file(upload["file_id"])
        source = parsers.FileDataSource(content, filename=upload["filename"], sheet_name=upload.get("sheet_name"))
        report_type = upload["report_type"]

        if upload.get("xero_native"):
            if report_type == "trial_balance":
                tb_current, tb_comparative = xero_reports.parse_trial_balance(source)
                data["tb_current"] = tb_current
                if not tb_comparative.empty:
                    data["tb_comparative"] = tb_comparative
            elif report_type == "nominal_activity":
                key = "nominal_current" if upload["period"] == "current" else "nominal_comparative"
                data[key] = xero_reports.parse_account_transactions(source)
            elif report_type == "aged_debtors":
                data["aged_debtors"] = xero_reports.parse_aged_report(source, "customer")
            elif report_type == "aged_creditors":
                data["aged_creditors"] = xero_reports.parse_aged_report(source, "supplier")
            continue

        if report_type in VAT_RECON_TYPES:
            # Every confirmed upload of these three types is concatenated,
            # not just the latest one under a single key - the VAT
            # Reconciliation workspace expects up to 10 filed-return files
            # (and a multi-sheet General Ledger) combined into one dataset,
            # each row tagged with the file it came from.
            df = parsers.apply_mapping(source, report_type, upload["mapping"] or {})
            df["source_file"] = upload.get("display_name") or upload["filename"]
            data.setdefault(report_type, []).append(df)
            continue

        if report_type in PAYE_RECON_TYPES:
            # Native BrightPay parse (see brightpay_reports.py), not
            # generic column-mapping - these were auto-confirmed on
            # upload the same way a Xero-native file is. Concatenated
            # across every confirmed upload of the type, same reasoning
            # as VAT_RECON_TYPES above: a client's tax year can span
            # more than one BrightPay export.
            if report_type == "paye_summary":
                df = brightpay_reports.parse_payroll_summary(source)
            elif report_type == "paye_p32":
                df = brightpay_reports.parse_p32(source)
            else:
                df = brightpay_reports.parse_pensions(source)
            data.setdefault(report_type, []).append(df)
            continue

        key = DATA_KEY.get((report_type, upload["period"]))
        if not key:
            continue
        df = parsers.apply_mapping(source, report_type, upload["mapping"] or {})
        data[key] = df

    for rt in VAT_RECON_TYPES + PAYE_RECON_TYPES:
        frames = data.get(rt)
        data[rt] = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

    # Xero's TB already categorises every account (Sales/Direct Costs/Overhead
    # = P&L, Bank/Current Asset/Current Liability/Equity/etc. = B/S), so P&L
    # and B/S can be derived when they weren't uploaded separately.
    if data.get("tb_current") is not None and not data["tb_current"].empty:
        if data.get("pl_current") is None or data.get("bs_current") is None:
            pl, bs = xero_reports.derive_pl_bs_from_tb(data["tb_current"])
            data.setdefault("pl_current", pl)
            data.setdefault("bs_current", bs)

    return data


def _period_days(job: dict) -> int:
    start, end = job.get("current_period_start"), job.get("current_period_end")
    if start and end:
        return (date.fromisoformat(end) - date.fromisoformat(start)).days + 1
    return 365


def _period_start(job: dict) -> date | None:
    start = job.get("current_period_start")
    return date.fromisoformat(start) if start else None


def _job_materiality(template: dict | None) -> tuple[float, float]:
    """(materiality, variance_pct_threshold) from a template's own config
    (storage.DEFAULT_TEMPLATE_CONFIG's "materiality" block), falling back
    to each check module's own constant when there's no template or no
    override set. Shared by the full Generate pipeline and every
    standalone per-section "Run" route (VAT/PAYE/Control Accounts/
    Debtors & Creditors/...), so a template's chosen threshold reaches a
    section whether it's tested on its own or as part of a full run."""
    materiality_cfg = template["config"].get("materiality", {}) if template else {}
    materiality = float(materiality_cfg.get("default_amount", recon.MATERIALITY_AMOUNT))
    variance_pct_threshold = float(materiality_cfg.get("variance_pct_threshold", recon.VARIANCE_PCT_THRESHOLD))
    return materiality, variance_pct_threshold


def _build_ct_computation(job: dict, data: dict, materiality: float = corporation_tax.MATERIALITY_AMOUNT) -> corporation_tax.CTComputation | None:
    pl = data.get("pl_current")
    if pl is None or pl.empty:
        return None
    accounting_profit = float(pl["amount"].sum())

    booked_tax_charge = None
    tb = data.get("tb_current")
    if tb is not None and not tb.empty:
        mask = tb["account_name"].str.lower().str.contains("corporation tax", na=False)
        if mask.any():
            booked_tax_charge = abs(float(tb.loc[mask, "balance"].sum()))

    return corporation_tax.compute(
        accounting_profit=accounting_profit,
        disallowable_additions=float(job.get("ct_disallowable_additions", 0.0)),
        capital_allowances=float(job.get("ct_capital_allowances", 0.0)),
        associated_companies=int(job.get("ct_associated_companies", 0)),
        period_days=_period_days(job),
        booked_tax_charge=booked_tax_charge,
        materiality=materiality,
    )


# Step labels shown on the live progress page (job_generate.html) - kept
# here as the single source of truth; the JS on that page has its own
# copy for offline rendering (a step's label needs to exist in the DOM
# before its first SSE event arrives), so keep the two in sync if this
# list ever changes.
GENERATE_STEPS = [
    "Loading and re-parsing confirmed uploads",
    "Running reconciliation checks",
    "Running anomaly detection",
    "Running compliance checks",
    "Generating AI-assisted reconciliation notes",
    "Building control account rollforwards",
    "Building the nominal activity matrix",
    "Computing Corporation Tax",
    "Building fixed asset registers",
    "Building the Excel workbook",
]

# Only steps 1 and 10 touch Postgres (loading uploads/template, saving the
# workbook and job) - the other eight run on data already in memory, so a
# retry there would never change a deterministic failure's outcome and
# they're left unwrapped. 3 attempts = 1 try + 2 retries; backoff grows
# with the attempt number so a real outage doesn't get hammered.
STEP_RETRY_ATTEMPTS = 3
STEP_RETRY_BASE_DELAY = 0.4


def _run_step_with_retry(event_fn, n: int, work):
    """Runs one step's work, retrying it with a short backoff before
    giving up. Yields "running" once, then a "retrying" event (attempt
    number + the error) for each failed attempt, then "done" and returns
    work()'s result - or re-raises the last error once every attempt is
    exhausted, exactly as an unwrapped step would (so a real, persistent
    failure still surfaces the same way it always has)."""
    yield event_fn(n, "running")
    attempt = 1
    while True:
        try:
            result = work()
        except Exception as exc:
            if attempt >= STEP_RETRY_ATTEMPTS:
                raise
            yield event_fn(n, "retrying", attempt=attempt, max_attempts=STEP_RETRY_ATTEMPTS, error=str(exc))
            time.sleep(STEP_RETRY_BASE_DELAY * attempt)
            attempt += 1
            continue
        break
    yield event_fn(n, "done")
    return result


def _generate_workbook_steps(job_id: str, job: dict, client: dict):
    """Does the real work of generating a working paper, one step at a
    time, yielding a small progress event before/after each of the ten
    steps in GENERATE_STEPS. This is the one place the generation logic
    lives - the classic POST route below just exhausts this generator
    without looking at the intermediate events (so its behaviour/response
    is unchanged from before this existed), and the SSE route streams
    those same events to the browser as they're produced. A step's real
    work happens between its "running" and "done" yields.

    Every event is also persisted onto job["progress"] (timestamped) as
    it fires, so the last run's step-by-step timing survives past the
    live stream - visible on the job page whether or not anyone watched
    it happen, and still there if the browser was closed mid-run.

    Steps 1 and 10 (the only two that touch Postgres) run through
    _run_step_with_retry, so a transient DB blip self-heals with a
    couple of quick retries instead of failing the whole run - see that
    function for the "retrying" events this can add between a step's
    "running" and "done"."""
    total = len(GENERATE_STEPS)
    job["progress"] = []
    storage.save_job(job)

    def record(evt: dict) -> dict:
        evt = {**evt, "at": datetime.now(timezone.utc).isoformat()}
        job["progress"].append(evt)
        storage.save_job(job)
        return evt

    def event(n: int, status: str, **extra) -> dict:
        return record({"step": n, "total": total, "label": GENERATE_STEPS[n - 1], "status": status, **extra})

    template = storage.get_template(client["practice_id"], client["template_id"]) if client.get("template_id") else None
    materiality, variance_pct_threshold = _job_materiality(template)

    def _step1():
        step_data = _load_canonical_data(job)
        step_pl_current = step_data.get("pl_current")
        step_profit = float(step_pl_current["amount"].sum()) if step_pl_current is not None and not step_pl_current.empty else None
        return step_data, step_pl_current, step_profit

    data, pl_current, current_year_profit = yield from _run_step_with_retry(event, 1, _step1)

    yield event(2, "running")
    results = recon.run_all_recons(data, materiality, variance_pct_threshold)
    vat_settings = vat_reconciliation.VatReconSettings(
        accounting_basis=job.get("vat_recon_basis", "accrual"),
        tolerance=job.get("vat_recon_tolerance", 0.0),
    )
    results = results + vat_reconciliation.reconcile(data, vat_settings)
    paye_settings = paye_reconciliation.PayeReconSettings(
        tolerance=job.get("paye_recon_tolerance", 0.0),
        date_window_days=job.get("paye_recon_date_window_days", paye_reconciliation.DEFAULT_DATE_WINDOW_DAYS),
    )
    results = results + paye_reconciliation.reconcile(data, paye_settings)
    yield event(2, "done")

    yield event(3, "running")
    results = results + anomaly_detection.run_all_anomaly_checks(data.get("nominal_current"))
    yield event(3, "done")

    yield event(4, "running")
    results = results + compliance_checks.run_all_compliance_checks(
        data.get("tb_current"), data.get("tb_comparative"), data.get("nominal_current"), current_year_profit,
    )
    period_end_str = job.get("current_period_end")
    results = results + [statutory_deadlines.build_result(date.fromisoformat(period_end_str) if period_end_str else None)]
    bs_statement_result = financial_statements.build_bs_statement(data.get("bs_current"), current_year_profit or 0.0, materiality)
    results = results + [going_concern.assess(bs_statement_result.statement, materiality)]
    results = results + [related_party_transactions.find_related_party_transactions(
        data.get("tb_current"), data.get("nominal_current"), materiality,
    )]
    yield event(4, "done")

    if template and template["config"].get("ai_reconciliation_notes", {}).get("enabled"):
        yield event(5, "running")
        _add_ai_reconciliation_notes(results, client, job)
        yield event(5, "done")
    else:
        yield event(5, "skipped", detail="not enabled for this template")

    yield event(6, "running")
    ca_results = control_accounts.build_all_rollforwards(
        data.get("tb_current"), data.get("tb_comparative"), data.get("nominal_current"),
        data.get("aged_debtors"), data.get("aged_creditors"), materiality,
    )
    results = results + [control_accounts.suggest_control_account_miscoding(
        data.get("tb_current"), data.get("nominal_current"),
        [(r.account_code, r.account_name) for r in ca_results],
        data.get("aged_debtors"), data.get("aged_creditors"), materiality,
    )]
    yield event(6, "done")

    yield event(7, "running")
    mx_results = nominal_matrix.build_all_matrices(data.get("tb_current"), data.get("nominal_current"), materiality=materiality)
    results = results + [nominal_matrix.suggest_unallocated_reallocations(
        data.get("nominal_current"), [r.account_code for r in mx_results] or None,
    )]
    yield event(7, "done")

    yield event(8, "running")
    ct_computation = _build_ct_computation(job, data, materiality)
    yield event(8, "done")

    yield event(9, "running")
    fixed_asset_result = fixed_assets.category_level_rollforward(
        data.get("tb_current"), data.get("tb_comparative"), data.get("nominal_current"),
        period_days=_period_days(job), materiality=materiality, period_start=_period_start(job),
    )
    asset_register_result = fixed_assets.asset_level_rollforward(
        data.get("fixed_asset_register"), data.get("nominal_current"), data.get("tb_current"),
        period_days=_period_days(job), materiality=materiality, period_start=_period_start(job),
    )
    results = results + [fixed_assets.suggest_capital_expenditure_reclassification(
        data.get("tb_current"), data.get("nominal_current"), data.get("fixed_asset_register"), threshold=materiality,
    )]
    results = results + [accruals_prepayments.build_schedule(
        data.get("tb_current"), data.get("tb_comparative"), data.get("nominal_current"), materiality,
    )]
    yield event(9, "done")

    def _step10():
        template_bytes = storage.load_file(template["file_id"]) if template else None
        if template_bytes is not None:
            wb = build_workbook_into_template(
                template_bytes, template["config"],
                job["client_name"], job["current_label"], job["comparative_label"], data, results,
                control_account_results=ca_results, matrix_results=mx_results, ct_computation=ct_computation,
                fixed_asset_result=fixed_asset_result, asset_register_result=asset_register_result,
                materiality=materiality, variance_pct_threshold=variance_pct_threshold,
            )
        else:
            wb = build_workbook(
                job["client_name"], job["current_label"], job["comparative_label"], data, results,
                control_account_results=ca_results, matrix_results=mx_results, ct_computation=ct_computation,
                fixed_asset_result=fixed_asset_result, asset_register_result=asset_register_result,
                materiality=materiality, variance_pct_threshold=variance_pct_threshold,
            )

        buffer = io.BytesIO()
        wb.save(buffer)
        output_filename = f"{job['client_name']} - Working Papers - {job['current_label']}.xlsx"
        # If a retry lands here (save_file succeeds, then save_job fails),
        # the prior attempt's file row is left in place, unreferenced -
        # harmless (never read back), and simpler than trying to reuse or
        # clean it up for a failure mode this rare.
        job["output_file_id"] = storage.save_file(
            "output", job_id, output_filename, buffer.getvalue(),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        job["status"] = "generated"
        job["summary"] = [{"name": r.name, "status": r.status, "message": r.message} for r in results]
        storage.save_job(job)

    yield from _run_step_with_retry(event, 10, _step10)

    yield record({"step": "complete", "redirect": f"/jobs/{job_id}"})


@app.post("/jobs/{job_id}/generate")
def generate(job_id: str, user: dict = Depends(auth.current_user_dep)):
    job, client = _authorize_job(user, job_id)
    for _event in _generate_workbook_steps(job_id, job, client):
        pass
    return RedirectResponse(f"/jobs/{job_id}", status_code=303)


@app.get("/jobs/{job_id}/generate/progress")
def generate_progress(request: Request, job_id: str, user: dict = Depends(auth.current_user_dep)):
    job, _client = _authorize_job(user, job_id)
    return templates.TemplateResponse("job_generate.html", {
        "request": request, "current_user": user, "job": job, "generate_steps": GENERATE_STEPS,
    })


@app.get("/jobs/{job_id}/generate/stream")
def generate_stream(job_id: str, user: dict = Depends(auth.current_user_dep)):
    """Server-Sent Events: the same ten-step generation as the POST route
    above, but each step's progress is pushed to the browser as it
    happens instead of only being visible once the whole thing finishes.
    A GET that triggers real work is unusual REST-wise, but it's what
    SSE/EventSource (the simplest way to stream progress without adding a
    JS framework or a job queue) requires, it's auth-gated the same as
    every other route, and it's only ever reached via the button on
    job_detail.html - never linked or crawlable."""
    job, client = _authorize_job(user, job_id)

    def event_source():
        try:
            for evt in _generate_workbook_steps(job_id, job, client):
                yield f"data: {json.dumps(evt)}\n\n"
        except Exception as exc:
            error_evt = {"step": "error", "message": str(exc), "at": datetime.now(timezone.utc).isoformat()}
            job["progress"].append(error_evt)
            storage.save_job(job)
            yield f"data: {json.dumps(error_evt)}\n\n"

    return StreamingResponse(
        event_source(), media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.get("/jobs/{job_id}/download")
def download(job_id: str, user: dict = Depends(auth.current_user_dep)):
    job, _client = _authorize_job(user, job_id)
    content = storage.load_file(job["output_file_id"])
    filename = f"{job['client_name']} - Working Papers - {job['current_label']}.xlsx"
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
