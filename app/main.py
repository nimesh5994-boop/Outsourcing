from pathlib import Path

from fastapi import FastAPI, Form, Request, UploadFile
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app import anomaly_detection, compliance_checks, control_accounts, corporation_tax, fixed_assets, mapping, nominal_matrix, parsers, recon, storage, xero_reports
from app.excel_builder import build_workbook, build_workbook_into_template
from app.models import PERIODS, PLATFORMS, REPORT_LABELS, REPORT_SCHEMAS, REPORT_TYPES, REQUIRED_FIELDS

# report types with a dedicated Xero report parser (no manual column mapping needed)
XERO_NATIVE_REPORT_TYPES = {"trial_balance", "nominal_activity", "aged_debtors", "aged_creditors"}

BASE_DIR = Path(__file__).resolve().parent
app = FastAPI(title="Working Paper Automation")
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

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
    practices = storage.list_practices()
    return templates.TemplateResponse("practices.html", {"request": request, "practices": practices})


@app.post("/practices")
def create_practice(name: str = Form(...)):
    practice = storage.create_practice(name.strip())
    return RedirectResponse(f"/practices/{practice['id']}", status_code=303)


@app.get("/practices/{practice_id}")
def practice_detail(request: Request, practice_id: str):
    practice = storage.get_practice(practice_id)
    templates_list = storage.list_templates(practice_id)
    return templates.TemplateResponse("practice_detail.html", {
        "request": request, "practice": practice, "templates_list": templates_list,
    })


@app.get("/practices/{practice_id}/clients")
def practice_clients(request: Request, practice_id: str):
    practice = storage.get_practice(practice_id)
    templates_list = storage.list_templates(practice_id)
    clients = storage.list_clients(practice_id)
    template_names = {t["id"]: t["name"] for t in templates_list}
    return templates.TemplateResponse("practice_clients.html", {
        "request": request, "practice": practice, "templates_list": templates_list,
        "clients": clients, "template_names": template_names,
    })


@app.post("/practices/{practice_id}/templates")
async def upload_template(practice_id: str, name: str = Form(...), file: UploadFile = None):
    content = await file.read()
    storage.create_template(practice_id, name.strip(), content, file.filename)
    return RedirectResponse(f"/practices/{practice_id}", status_code=303)


@app.get("/practices/{practice_id}/templates/{template_id}")
def template_detail(request: Request, practice_id: str, template_id: str):
    import json as _json
    template = storage.get_template(practice_id, template_id)
    return templates.TemplateResponse("template_detail.html", {
        "request": request, "practice_id": practice_id, "template": template,
        "config_json": _json.dumps(template["config"], indent=2),
    })


@app.post("/practices/{practice_id}/templates/{template_id}/config")
async def save_template_config(request: Request, practice_id: str, template_id: str):
    import json as _json
    form = await request.form()
    template = storage.get_template(practice_id, template_id)
    try:
        template["config"] = _json.loads(form.get("config_json"))
    except _json.JSONDecodeError as exc:
        return templates.TemplateResponse("template_detail.html", {
            "request": request, "practice_id": practice_id, "template": template,
            "config_json": form.get("config_json"), "error": f"Invalid JSON: {exc}",
        })
    storage.save_template(template)
    return RedirectResponse(f"/practices/{practice_id}/templates/{template_id}", status_code=303)


@app.post("/practices/{practice_id}/templates/{template_id}/make-default")
def make_template_default(practice_id: str, template_id: str):
    practice = storage.get_practice(practice_id)
    practice["default_template_id"] = template_id
    storage.save_practice(practice)
    return RedirectResponse(f"/practices/{practice_id}", status_code=303)


@app.post("/practices/{practice_id}/clients")
def create_client(practice_id: str, name: str = Form(...), template_id: str = Form("")):
    practice = storage.get_practice(practice_id)
    chosen_template = template_id or practice.get("default_template_id")
    client = storage.create_client(practice_id, name.strip(), chosen_template)
    return RedirectResponse(f"/clients/{client['id']}", status_code=303)


@app.get("/clients/{client_id}")
def client_detail(request: Request, client_id: str):
    client = storage.get_client(client_id)
    jobs = storage.list_jobs(client_id)
    return templates.TemplateResponse("client_detail.html", {"request": request, "client": client, "jobs": jobs})


@app.post("/clients/{client_id}/jobs")
def create_job(
    client_id: str,
    current_period_start: str = Form(...), current_period_end: str = Form(...),
    comparative_period_start: str = Form(""), comparative_period_end: str = Form(""),
):
    client = storage.get_client(client_id)
    job = storage.create_job(
        client_id, client["name"],
        current_period_start, current_period_end,
        comparative_period_start or None, comparative_period_end or None,
    )
    return RedirectResponse(f"/jobs/{job['id']}", status_code=303)


@app.get("/jobs/{job_id}")
def job_detail(request: Request, job_id: str):
    job = storage.get_job(job_id)
    return templates.TemplateResponse("job_detail.html", {
        "request": request, "job": job,
        "report_types": REPORT_TYPES, "report_labels": REPORT_LABELS,
        "platforms": PLATFORMS, "periods": PERIODS,
    })


@app.post("/jobs/{job_id}/tax-inputs")
def save_tax_inputs(
    job_id: str,
    ct_associated_companies: int = Form(0),
    ct_disallowable_additions: float = Form(0.0),
    ct_capital_allowances: float = Form(0.0),
):
    job = storage.get_job(job_id)
    job["ct_associated_companies"] = ct_associated_companies
    job["ct_disallowable_additions"] = ct_disallowable_additions
    job["ct_capital_allowances"] = ct_capital_allowances
    storage.save_job(job)
    return RedirectResponse(f"/jobs/{job_id}", status_code=303)


@app.post("/jobs/{job_id}/uploads")
async def upload_file(job_id: str, report_type: str = Form(...), period: str = Form(...),
                       platform: str = Form(...), file: UploadFile = None):
    job = storage.get_job(job_id)
    dest_dir = storage.uploads_dir(job_id)
    saved_path = dest_dir / f"{report_type}__{period}__{file.filename}"
    content = await file.read()
    saved_path.write_bytes(content)

    source = parsers.FileDataSource(saved_path)
    expected_end = _expected_period_end(job, period)

    # Xero exports have a known, specific layout (grouped reports, embedded
    # comparative TB) that a dedicated parser handles directly - no manual
    # column mapping step needed. Falls back to generic mapping if the file
    # doesn't actually match the expected Xero layout (a real signal that
    # either the wrong report type was picked, or it isn't a Xero export).
    validation_note = ""
    period_check = {"status": "unknown", "message": ""}
    if platform == "xero" and report_type in XERO_NATIVE_REPORT_TYPES:
        try:
            _validate_xero_parse(report_type, source)
        except Exception as exc:
            validation_note = (
                f"This file doesn't look like a standard Xero {REPORT_LABELS[report_type]} export "
                f"({exc}) - check you selected the right report type, or map its columns manually below."
            )
        else:
            period_check = xero_reports.check_period(xero_reports.extract_period_info(source), expected_end)
            columns = source.raw_columns()
            upload_id = storage.add_upload(job, report_type, period, platform, file.filename, str(saved_path), columns)
            job = storage.get_job(job_id)
            job["uploads"][upload_id]["mapping"] = {}
            job["uploads"][upload_id]["confirmed"] = True
            job["uploads"][upload_id]["xero_native"] = True
            job["uploads"][upload_id]["period_check"] = period_check
            storage.save_job(job)
            return RedirectResponse(f"/jobs/{job_id}", status_code=303)

    columns = source.raw_columns()
    upload_id = storage.add_upload(job, report_type, period, platform, file.filename, str(saved_path), columns)

    # try a saved profile first, fall back to alias-based suggestion
    profile = mapping.load_profile(job["client_id"], report_type, platform)
    if profile:
        suggestion = {col: profile.get(col) for col in columns}
    else:
        suggestion = mapping.suggest_mapping(report_type, columns)

    # simple category check: did we manage to guess a mapping for the fields
    # this report type actually needs? If none matched at all, the file is
    # probably not what was selected in the report-type dropdown.
    required = REQUIRED_FIELDS.get(report_type, [])
    matched_required = {v for v in suggestion.values() if v} & set(required)
    if not validation_note and required and not matched_required:
        validation_note = (
            f"None of this file's columns look like a {REPORT_LABELS[report_type]} - "
            f"double check you selected the right report type before mapping columns."
        )

    job = storage.get_job(job_id)
    job["uploads"][upload_id]["mapping"] = suggestion
    job["uploads"][upload_id]["validation_note"] = validation_note
    storage.save_job(job)

    return RedirectResponse(f"/jobs/{job_id}/uploads/{upload_id}/mapping", status_code=303)


def _expected_period_end(job: dict, period: str):
    from datetime import date
    key = "current_period_end" if period == "current" else "comparative_period_end"
    value = job.get(key)
    return date.fromisoformat(value) if value else None


def _validate_xero_parse(report_type: str, source: parsers.DataSource) -> None:
    if report_type == "trial_balance":
        xero_reports.parse_trial_balance(source)
    elif report_type == "nominal_activity":
        xero_reports.parse_account_transactions(source)
    elif report_type == "aged_debtors":
        xero_reports.parse_aged_report(source, "customer")
    elif report_type == "aged_creditors":
        xero_reports.parse_aged_report(source, "supplier")


@app.get("/jobs/{job_id}/uploads/{upload_id}/mapping")
def mapping_form(request: Request, job_id: str, upload_id: str):
    job = storage.get_job(job_id)
    upload = job["uploads"][upload_id]
    canonical_fields = REPORT_SCHEMAS[upload["report_type"]]
    return templates.TemplateResponse("mapping.html", {
        "request": request, "job": job, "upload": upload,
        "canonical_fields": canonical_fields, "report_label": REPORT_LABELS[upload["report_type"]],
    })


@app.post("/jobs/{job_id}/uploads/{upload_id}/mapping")
async def save_mapping(request: Request, job_id: str, upload_id: str):
    form = await request.form()
    job = storage.get_job(job_id)
    upload = job["uploads"][upload_id]

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

    return RedirectResponse(f"/jobs/{job_id}", status_code=303)


def _load_canonical_data(job: dict) -> dict:
    data = {}
    for upload in job["uploads"].values():
        if not upload["confirmed"]:
            continue
        source = parsers.FileDataSource(upload["path"])
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
def generate(job_id: str):
    job = storage.get_job(job_id)
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

    client = storage.get_client(job["client_id"])
    template = storage.get_template(client["practice_id"], client["template_id"]) if client and client.get("template_id") else None

    if template and Path(template["file_path"]).exists():
        wb = build_workbook_into_template(
            template["file_path"], template["config"],
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

    out_path = storage.output_dir(job_id) / "working_paper.xlsx"
    wb.save(out_path)

    job["status"] = "generated"
    job["summary"] = [{"name": r.name, "status": r.status, "message": r.message} for r in results]
    storage.save_job(job)

    return RedirectResponse(f"/jobs/{job_id}", status_code=303)


@app.get("/jobs/{job_id}/download")
def download(job_id: str):
    job = storage.get_job(job_id)
    out_path = storage.output_dir(job_id) / "working_paper.xlsx"
    filename = f"{job['client_name']} - Working Papers - {job['current_label']}.xlsx"
    return FileResponse(out_path, filename=filename,
                         media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
