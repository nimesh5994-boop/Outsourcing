"""Orchestrates a pull from Xero into a job: the TB/account metadata in one
quick pass, and the ledger (Journals) in resumable batches.

The pulled data is stored as job files (so deleting a job removes it) and
described by job["xero_pull"]. main.py's _load_canonical_data reads it back
as the base layer - any confirmed file upload for the same report then
overrides it, so a preparer can always correct or replace a pulled report
by uploading an export.
"""
import json
import time
from datetime import datetime, timezone

import pandas as pd

from app import storage, xero_api

JOURNAL_PAGE_SIZE = 100
DEFAULT_BATCH_SECONDS = 35


def _save_json(job_id: str, name: str, payload) -> str:
    return storage.save_file("xero_raw", job_id, f"{name}.json", json.dumps(payload).encode(), content_type="application/json")


def _save_frame(job_id: str, name: str, df: pd.DataFrame) -> str:
    return storage.save_file("xero_snapshot", job_id, f"{name}.json", xero_api.frame_to_bytes(df), content_type="application/json")


def _discard_files(file_ids) -> None:
    for fid in file_ids:
        storage.delete_file(fid)


def pull_core(job: dict, client: xero_api.XeroClient, tenant_name: str) -> dict:
    """Trial Balance (job's current and, if set, comparative period end) +
    the account list that gives each TB row Xero's own account type. Raises
    XeroError subclasses untouched - the route decides what to tell the user.
    Replaces any earlier pull for this job (old files removed only after the
    new ones are saved)."""
    previous = job.get("xero_pull") or {}

    accounts_raw = client.get("Accounts")
    organisation_raw = client.get("Organisation")
    accounts = xero_api.index_accounts(accounts_raw)

    tb_current_raw = client.get("Reports/TrialBalance", {"date": job["current_period_end"]})
    tb_current = xero_api.trial_balance_frame(tb_current_raw, accounts)

    files = {"tb_current": _save_frame(job["id"], "tb_current", tb_current)}
    raw = {"tb_current": _save_json(job["id"], "raw_tb_current", tb_current_raw),
           "accounts": _save_json(job["id"], "raw_accounts", accounts_raw),
           "organisation": _save_json(job["id"], "raw_organisation", organisation_raw)}
    warnings = []
    tb_comparative = None

    fy_warning = xero_api.financial_year_warning(organisation_raw, job["current_period_start"], job["current_period_end"])
    if fy_warning:
        warnings.append(fy_warning)

    if job.get("comparative_period_end"):
        tb_comp_raw = client.get("Reports/TrialBalance", {"date": job["comparative_period_end"]})
        tb_comparative = xero_api.trial_balance_frame(tb_comp_raw, accounts)
        files["tb_comparative"] = _save_frame(job["id"], "tb_comparative", tb_comparative)
        raw["tb_comparative"] = _save_json(job["id"], "raw_tb_comparative", tb_comp_raw)

    untyped = int((tb_current["account_type"] == "").sum())
    if untyped:
        warnings.append(f"{untyped} Trial Balance account(s) couldn't be matched to Xero's account list, so they have no account type "
                        "- they will be treated as balance sheet items. Check the pulled TB before relying on the P&L/Balance Sheet split.")

    journals_state = {"status": "not_started"}
    if not client.has_scope(xero_api.JOURNALS_SCOPE):
        journals_state = {"status": "unavailable", "note": "This connection wasn't granted ledger (Journals) access - nominal-ledger "
                          "schedules still need an uploaded Account Transactions export."}

    job["xero_pull"] = {
        "pulled_at": datetime.now(timezone.utc).isoformat(), "tenant_name": tenant_name, "warnings": warnings,
        "files": files, "raw": raw, "journals": journals_state,
        "summary": {"tb_current_rows": len(tb_current), "tb_comparative_rows": None if tb_comparative is None else len(tb_comparative)},
    }
    storage.save_job(job)
    _discard_files(list((previous.get("files") or {}).values()) + list((previous.get("raw") or {}).values())
                   + list((previous.get("journals") or {}).get("batch_file_ids", [])))
    return job["xero_pull"]


def pull_journals_batch(job: dict, client: xero_api.XeroClient, budget_seconds: float = DEFAULT_BATCH_SECONDS,
                        clock=time.monotonic) -> dict:
    """Fetches journals page by page until the time budget is spent or
    Xero has no more, keeping only journals inside the job's current
    period (the endpoint can't filter by date - it can only resume after a
    journal number - so earlier history is paged past, not kept). Safe to
    call repeatedly: progress (next offset, kept journals) lives on the job.
    On the final batch, builds nominal_current and clears the scratch files."""
    pull = job.get("xero_pull")
    if not pull:
        raise xero_api.XeroError("Pull the Trial Balance first.")
    state = pull["journals"]
    if state["status"] in ("unavailable", "complete"):
        return state

    start, end = job["current_period_start"], job["current_period_end"]
    offset = state.get("next_offset", 0)
    batch_files = list(state.get("batch_file_ids", []))
    kept_total = state.get("kept", 0)
    deadline = clock() + budget_seconds
    kept_this_run, finished = [], False

    try:
        while clock() < deadline:
            page = client.get("Journals", {"offset": offset}).get("Journals", [])
            if not page:
                finished = True
                break
            offset = max(j.get("JournalNumber", 0) for j in page)
            in_range = [j for j in page if _in_period(j, start, end)]
            kept_this_run.extend(in_range)
            if len(page) < JOURNAL_PAGE_SIZE:
                finished = True
                break
    except xero_api.XeroForbidden:
        pull["journals"] = {"status": "unavailable", "note": "Xero refused ledger (Journals) access for this app/connection - nominal-ledger "
                            "schedules still need an uploaded Account Transactions export."}
        _discard_files(batch_files)
        storage.save_job(job)
        return pull["journals"]

    if kept_this_run:
        batch_files.append(_save_json(job["id"], f"journals_batch_{len(batch_files) + 1}", kept_this_run))
        kept_total += len(kept_this_run)

    if not finished:
        pull["journals"] = {"status": "in_progress", "next_offset": offset, "batch_file_ids": batch_files, "kept": kept_total}
        storage.save_job(job)
        return pull["journals"]

    journals = []
    for fid in batch_files:
        journals.extend(json.loads(storage.load_file(fid)))
    nominal = xero_api.journals_frame(journals, start, end)
    pull["files"]["nominal_current"] = _save_frame(job["id"], "nominal_current", nominal)
    _discard_files(batch_files)
    pull["journals"] = {"status": "complete", "kept": kept_total, "lines": int(len(nominal))}
    storage.save_job(job)
    return pull["journals"]


def _in_period(journal: dict, start: str, end: str) -> bool:
    d = xero_api._journal_date(journal.get("JournalDate"))
    return pd.notna(d) and pd.Timestamp(start) <= d <= pd.Timestamp(end)


def load_snapshot(job: dict) -> dict:
    """Pulled frames keyed like _load_canonical_data's own output
    (tb_current / tb_comparative / nominal_current)."""
    pull = job.get("xero_pull") or {}
    out = {}
    for key, file_id in (pull.get("files") or {}).items():
        blob = storage.load_file(file_id)
        if blob is not None:
            out[key] = xero_api.frame_from_bytes(blob)
    return out
