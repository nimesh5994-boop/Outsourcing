import io
from pathlib import Path

from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app import anomaly_detection, auth, compliance_checks, control_accounts, corporation_tax, document_detection, fixed_assets, mapping, nominal_matrix, parsers, recon, storage, xero_reports
from app.excel_builder import build_workbook, build_workbook_into_template
from app.models import PERIODS, PLATFORMS, REPORT_LABELS, REPORT_SCHEMAS, REPORT_TYPES, REQUIRED_FIELDS

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
}


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
    _authorize_practice(user, practice_id)
    auth.require_role(user, "partner", "manager")
    template = storage.get_template(practice_id, template_id)
    return templates.TemplateResponse("template_detail.html", {
        "request": request, "current_user": user, "practice_id": practice_id, "template": template,
        "config_json": _json.dumps(template["config"], indent=2),
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
    _authorize_practice(user, practice_id)
    auth.require_role(user, "partner")
    practice_users = storage.list_users(practice_id)
    clients = storage.list_clients(practice_id)
    access_by_user = {u["id"]: set(storage.list_client_access(u["id"])) for u in practice_users if u["role"] == "preparer"}
    return templates.TemplateResponse("users.html", {
        "request": request, "current_user": user, "practice_id": practice_id,
        "practice_users": practice_users, "clients": clients, "access_by_user": access_by_user,
        "roles": auth.ROLES,
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
    return templates.TemplateResponse("client_detail.html", {
        "request": request, "current_user": user, "client": client, "jobs": jobs,
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
    type_hint = section_report_type if section_report_type in REPORT_TYPES else None

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

        key = DATA_KEY.get((report_type, upload["period"]))
        if not key:
            continue
        df = parsers.apply_mapping(source, report_type, upload["mapping"] or {})
        data[key] = df

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
        from datetime import date
        return (date.fromisoformat(end) - date.fromisoformat(start)).days + 1
    return 365


def _build_ct_computation(job: dict, data: dict) -> corporation_tax.CTComputation | None:
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
    )


@app.post("/jobs/{job_id}/generate")
def generate(job_id: str, user: dict = Depends(auth.current_user_dep)):
    job, client = _authorize_job(user, job_id)
    data = _load_canonical_data(job)
    pl_current = data.get("pl_current")
    current_year_profit = float(pl_current["amount"].sum()) if pl_current is not None and not pl_current.empty else None
    results = (
        recon.run_all_recons(data)
        + anomaly_detection.run_all_anomaly_checks(data.get("nominal_current"))
        + compliance_checks.run_all_compliance_checks(
            data.get("tb_current"), data.get("tb_comparative"), data.get("nominal_current"), current_year_profit,
        )
    )

    ca_results = control_accounts.build_all_rollforwards(
        data.get("tb_current"), data.get("tb_comparative"), data.get("nominal_current"),
        data.get("aged_debtors"), data.get("aged_creditors"),
    )
    mx_results = nominal_matrix.build_all_matrices(data.get("tb_current"), data.get("nominal_current"))
    ct_computation = _build_ct_computation(job, data)

    fixed_asset_result = fixed_assets.category_level_rollforward(
        data.get("tb_current"), data.get("tb_comparative"), data.get("nominal_current"),
    )
    asset_register_result = fixed_assets.asset_level_rollforward(
        data.get("fixed_asset_register"), data.get("nominal_current"), data.get("tb_current"),
        period_days=_period_days(job),
    )

    template = storage.get_template(client["practice_id"], client["template_id"]) if client.get("template_id") else None
    template_bytes = storage.load_file(template["file_id"]) if template else None

    if template_bytes is not None:
        wb = build_workbook_into_template(
            template_bytes, template["config"],
            job["client_name"], job["current_label"], job["comparative_label"], data, results,
            control_account_results=ca_results, matrix_results=mx_results, ct_computation=ct_computation,
            fixed_asset_result=fixed_asset_result, asset_register_result=asset_register_result,
        )
    else:
        wb = build_workbook(
            job["client_name"], job["current_label"], job["comparative_label"], data, results,
            control_account_results=ca_results, matrix_results=mx_results, ct_computation=ct_computation,
            fixed_asset_result=fixed_asset_result, asset_register_result=asset_register_result,
        )

    buffer = io.BytesIO()
    wb.save(buffer)
    output_filename = f"{job['client_name']} - Working Papers - {job['current_label']}.xlsx"
    job["output_file_id"] = storage.save_file(
        "output", job_id, output_filename, buffer.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    job["status"] = "generated"
    job["summary"] = [{"name": r.name, "status": r.status, "message": r.message} for r in results]
    storage.save_job(job)

    return RedirectResponse(f"/jobs/{job_id}", status_code=303)


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
