"""Direct Xero API access: OAuth 2.0 (authorization code flow), a small
rate-limit-aware client, and converters from Xero's API JSON into the same
canonical DataFrames xero_reports.py builds from an exported file - so a
pulled dataset runs through the rest of the pipeline (reconciliations,
formula-linked schedules, ETB) completely unchanged.

Xero has no static "API key" for this: an app registered in the Xero
developer portal has a client id + secret, and each client organisation is
connected by a user consenting once (access tokens last 30 minutes, the
rotating single-use refresh token is how access is kept alive afterwards).

What is pulled, and why only this: the Trial Balance report (current and
comparative year), Accounts (for Xero's own account type per code - the TB
report carries only a presentation grouping), Organisation (financial year
end, to catch a job period the year-to-date TB can't represent), and - only
when the app was granted it - Journals, the one endpoint that exposes every
posting and so stands in for the "Account Transactions" export. Xero gates
Journals behind its top developer tier plus a security assessment, so a
missing/denied scope is a normal, handled outcome (the TB-driven schedules
still work; nominal-ledger schedules keep using an uploaded export).
"""
import base64
import calendar
import hashlib
import json
import os
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone

import pandas as pd
from cryptography.fernet import Fernet, InvalidToken

AUTHORIZE_URL = "https://login.xero.com/identity/connect/authorize"
TOKEN_URL = "https://identity.xero.com/connect/token"
CONNECTIONS_URL = "https://api.xero.com/connections"
API_BASE = "https://api.xero.com/api.xro/2.0"

# Granular scopes (apps registered from 2 March 2026 can't use the broad
# accounting.reports.read at all). Journals has no confirmed granular
# scope name for new apps, so it is opt-in via XERO_SCOPES rather than
# requested by default - requesting a scope the app isn't entitled to
# makes Xero reject the whole authorization, taking the TB pull down too.
DEFAULT_SCOPES = "offline_access accounting.settings accounting.reports.trialbalance.read"
JOURNALS_SCOPE = "accounting.journals.read"

MAX_RATE_LIMIT_WAITS = 3
MAX_RETRY_AFTER_SECONDS = 20
TOKEN_REFRESH_MARGIN_SECONDS = 60


class XeroError(Exception):
    pass


class XeroNotConfigured(XeroError):
    pass


class XeroAuthError(XeroError):
    """Refresh failed or the user revoked access - the client must be reconnected."""


class XeroForbidden(XeroError):
    """The app lacks the scope/tier for this endpoint."""


def configured() -> bool:
    return bool(os.environ.get("XERO_CLIENT_ID") and os.environ.get("XERO_CLIENT_SECRET"))


def _credentials() -> tuple[str, str]:
    cid, secret = os.environ.get("XERO_CLIENT_ID"), os.environ.get("XERO_CLIENT_SECRET")
    if not cid or not secret:
        raise XeroNotConfigured("XERO_CLIENT_ID / XERO_CLIENT_SECRET are not set.")
    return cid, secret


def requested_scopes() -> str:
    return os.environ.get("XERO_SCOPES", "").strip() or DEFAULT_SCOPES


# ---------- token encryption (refresh tokens are long-lived credentials) ----------

def _fernet() -> Fernet:
    key = hashlib.sha256(("xero-token-encryption:" + os.environ["SECRET_KEY"]).encode()).digest()
    return Fernet(base64.urlsafe_b64encode(key))


def encrypt_token(token: str) -> str:
    return _fernet().encrypt(token.encode()).decode()


def decrypt_token(blob: str) -> str:
    try:
        return _fernet().decrypt(blob.encode()).decode()
    except InvalidToken as exc:
        raise XeroAuthError("Stored Xero credentials can't be decrypted (SECRET_KEY changed?) - reconnect Xero.") from exc


# ---------- HTTP ----------

def default_http(method: str, url: str, headers: dict, body: bytes | None = None) -> tuple[int, dict, bytes]:
    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.status, {k.lower(): v for k, v in resp.headers.items()}, resp.read()
    except urllib.error.HTTPError as exc:
        return exc.code, {k.lower(): v for k, v in exc.headers.items()}, exc.read()


def _basic_auth_header() -> str:
    cid, secret = _credentials()
    return "Basic " + base64.b64encode(f"{cid}:{secret}".encode()).decode()


def authorize_url(redirect_uri: str, state: str) -> str:
    cid, _ = _credentials()
    return AUTHORIZE_URL + "?" + urllib.parse.urlencode({
        "response_type": "code", "client_id": cid, "redirect_uri": redirect_uri,
        "scope": requested_scopes(), "state": state,
    })


def _token_request(form: dict, http=None) -> dict:
    http = http or default_http
    status, _, body = http(
        "POST", TOKEN_URL,
        {"Authorization": _basic_auth_header(), "Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"},
        urllib.parse.urlencode(form).encode(),
    )
    payload = json.loads(body or b"{}")
    if status != 200 or "access_token" not in payload:
        raise XeroAuthError(f"Xero token request failed ({status}): {payload.get('error_description') or payload.get('error') or 'unknown error'}")
    return payload


def exchange_code(code: str, redirect_uri: str, http=None, now=time.time) -> dict:
    payload = _token_request({"grant_type": "authorization_code", "code": code, "redirect_uri": redirect_uri}, http)
    return _token_fields(payload, now)


def refresh(refresh_token: str, http=None, now=time.time) -> dict:
    payload = _token_request({"grant_type": "refresh_token", "refresh_token": refresh_token}, http)
    return _token_fields(payload, now)


def _token_fields(payload: dict, now) -> dict:
    return {
        "access_token_enc": encrypt_token(payload["access_token"]),
        "refresh_token_enc": encrypt_token(payload["refresh_token"]),
        "expires_at": now() + int(payload.get("expires_in", 1800)),
        "scopes": payload.get("scope", ""),
    }


def list_tenants(access_token: str, http=None) -> list[dict]:
    http = http or default_http
    status, _, body = http("GET", CONNECTIONS_URL, {"Authorization": f"Bearer {access_token}", "Accept": "application/json"})
    if status != 200:
        raise XeroError(f"Couldn't list Xero organisations ({status}).")
    return [{"connection_id": c.get("id", ""), "tenant_id": c["tenantId"], "tenant_name": c.get("tenantName", "")}
            for c in json.loads(body) if c.get("tenantType", "ORGANISATION") == "ORGANISATION"]


def disconnect(connection_id: str, access_token: str, http=None) -> None:
    """Revokes this app's access to one organisation on Xero's side (best
    effort - the caller deletes the stored credentials regardless)."""
    (http or default_http)("DELETE", f"{CONNECTIONS_URL}/{connection_id}", {"Authorization": f"Bearer {access_token}"})


class XeroClient:
    """One connected organisation. `connection` is the stored dict (see
    main.py's xero_connection entity); `save` persists it whenever a token
    refresh rotates the credentials - the refresh token is single-use, so a
    refresh whose result isn't saved would lock the connection out."""

    def __init__(self, connection: dict, save, http=None, sleep=time.sleep, now=time.time):
        self.connection, self._save, self._http, self._sleep, self._now = connection, save, http, sleep, now

    def has_scope(self, scope: str) -> bool:
        return scope in (self.connection.get("scopes") or "").split()

    def _access_token(self, force_refresh: bool = False) -> str:
        c = self.connection
        if force_refresh or self._now() >= c.get("expires_at", 0) - TOKEN_REFRESH_MARGIN_SECONDS:
            c.update(refresh(decrypt_token(c["refresh_token_enc"]), self._http, self._now))
            self._save(c)
        return decrypt_token(c["access_token_enc"])

    def get(self, path: str, params: dict | None = None) -> dict:
        url = f"{API_BASE}/{path}" + (("?" + urllib.parse.urlencode(params)) if params else "")
        refreshed_after_401 = False
        waits = 0
        while True:
            headers = {
                "Authorization": f"Bearer {self._access_token()}", "xero-tenant-id": self.connection["tenant_id"],
                "Accept": "application/json",
            }
            status, resp_headers, body = (self._http or default_http)("GET", url, headers)
            if status == 200:
                return json.loads(body)
            if status == 429 and waits < MAX_RATE_LIMIT_WAITS:
                waits += 1
                self._sleep(min(float(resp_headers.get("retry-after", 5)), MAX_RETRY_AFTER_SECONDS))
                continue
            if status == 401 and not refreshed_after_401:
                refreshed_after_401 = True
                self._access_token(force_refresh=True)
                continue
            if status in (401, 403):
                detail = body.decode("utf-8", "replace")[:300]
                if "insufficient_scope" in detail.lower() or status == 403:
                    raise XeroForbidden(f"Xero refused {path} - this app lacks the scope or tier for it ({status}).")
                raise XeroAuthError("Xero rejected the stored credentials - reconnect Xero.")
            raise XeroError(f"Xero returned {status} for {path}: {body.decode('utf-8', 'replace')[:300]}")


# ---------- Xero's account type tokens -> the labels its report exports use ----------
# The rest of the app keys off the export's display labels (e.g. "Direct
# Costs", "Non-current Liability", "Inventory"), so API tokens are
# translated to those exact strings.
ACCOUNT_TYPE_LABELS = {
    "BANK": "Bank", "CURRENT": "Current Asset", "CURRLIAB": "Current Liability", "DEPRECIATN": "Depreciation",
    "DIRECTCOSTS": "Direct Costs", "EQUITY": "Equity", "EXPENSE": "Expense", "FIXED": "Fixed Asset",
    "INVENTORY": "Inventory", "LIABILITY": "Liability", "NONCURRENT": "Non-current Asset",
    "OTHERINCOME": "Other Income", "OVERHEADS": "Overhead", "PREPAYMENT": "Prepayment", "REVENUE": "Revenue",
    "SALES": "Sales", "TERMLIAB": "Non-current Liability", "PAYGLIABILITY": "Current Liability",
    "SUPERANNUATIONEXPENSE": "Expense", "SUPERANNUATIONLIABILITY": "Current Liability",
    "WAGESEXPENSE": "Expense", "WAGESPAYABLELIABILITY": "Current Liability",
}


def account_type_label(token: str) -> str:
    token = (token or "").strip()
    return ACCOUNT_TYPE_LABELS.get(token.upper(), token.title())


def index_accounts(accounts_response: dict) -> dict:
    """AccountID -> {code, name, type}."""
    return {
        a["AccountID"]: {"code": a.get("Code", "") or "", "name": a.get("Name", ""), "type": account_type_label(a.get("Type", ""))}
        for a in accounts_response.get("Accounts", [])
    }


# ---------- Trial Balance report -> tb DataFrame ----------

_CODE_SUFFIX = re.compile(r"^(.*?)\s*\(([^()]*)\)\s*$")


def _num(value) -> float:
    if value is None or value == "":
        return 0.0
    try:
        return float(str(value).replace(",", ""))
    except ValueError:
        return 0.0


def _walk_rows(rows: list):
    for row in rows:
        yield row
        if row.get("Rows"):
            yield from _walk_rows(row["Rows"])


def trial_balance_frame(report_response: dict, accounts_by_id: dict) -> pd.DataFrame:
    """Same columns/semantics as xero_reports.parse_trial_balance's
    tb_current: year-to-date debit/credit per account (the exported TB's
    own 'Debit/Credit - Year to date'), account_type as Xero's own label,
    and - matching that parser - a name-only account (e.g. a bank feed with
    no nominal code) keyed by its name."""
    reports = report_response.get("Reports") or []
    if not reports:
        raise XeroError("Xero returned no Trial Balance report.")

    columns: dict[str, int] = {}
    out_rows = []
    for row in _walk_rows(reports[0].get("Rows", [])):
        row_type = row.get("RowType")
        cells = row.get("Cells") or []
        if row_type == "Header" and not columns:
            columns = {str(c.get("Value", "")).strip().lower(): i for i, c in enumerate(cells)}
            continue
        if row_type != "Row" or not cells:
            continue

        label = str(cells[0].get("Value", "")).strip()
        account_id = next((a.get("Value") for a in cells[0].get("Attributes", []) if a.get("Id") == "account"), None)
        known = accounts_by_id.get(account_id) if account_id else None
        if known:
            code, name, account_type = known["code"], known["name"], known["type"]
        else:
            m = _CODE_SUFFIX.match(label)
            code, name, account_type = (m.group(2), m.group(1), "") if m else ("", label, "")
        code = code or name

        debit_col = columns.get("ytd debit", columns.get("debit"))
        credit_col = columns.get("ytd credit", columns.get("credit"))
        if debit_col is None or credit_col is None:
            raise XeroError("Xero's Trial Balance has no Debit/Credit columns in the expected places - the report layout may have changed.")
        out_rows.append({
            "account_code": code, "account_name": name, "account_type": account_type,
            "debit": _num(cells[debit_col].get("Value")), "credit": _num(cells[credit_col].get("Value")),
        })

    if not out_rows:
        raise XeroError("Xero's Trial Balance had no account rows.")
    df = pd.DataFrame(out_rows)
    df["balance"] = df["debit"] - df["credit"]
    return df


# ---------- Journals -> nominal activity DataFrame ----------

SOURCE_TYPE_LABELS = {
    "ACCREC": "Receivable Invoice", "ACCPAY": "Payable Invoice", "ACCRECCREDIT": "Receivable Credit Note",
    "ACCPAYCREDIT": "Payable Credit Note", "ACCRECPAYMENT": "Receivable Payment", "ACCPAYPAYMENT": "Payable Payment",
    "ARCREDITPAYMENT": "Receivable Credit Note Payment", "APCREDITPAYMENT": "Payable Credit Note Payment",
    "CASHREC": "Receive Money", "CASHPAY": "Spend Money", "TRANSFER": "Bank Transfer",
    "ARPREPAYMENT": "Receivable Prepayment", "APPREPAYMENT": "Payable Prepayment",
    "AROVERPAYMENT": "Receivable Overpayment", "APOVERPAYMENT": "Payable Overpayment",
    "EXPCLAIM": "Expense Claim", "EXPPAYMENT": "Expense Claim Payment", "MANJOURNAL": "Manual Journal",
    "PAYSLIP": "Payslip", "WAGEPAYABLE": "Wage Payable",
}

_MS_DATE = re.compile(r"/Date\((-?\d+)")

NOMINAL_COLUMNS = ["date", "account_code", "account_name", "reference", "description", "contact", "source_type",
                   "debit", "credit", "vat_amount", "contra_code", "contra_name", "contra_needs_review", "net"]


def _journal_date(value) -> pd.Timestamp:
    m = _MS_DATE.match(str(value))
    if m:
        return pd.Timestamp(datetime.fromtimestamp(int(m.group(1)) / 1000, tz=timezone.utc).replace(tzinfo=None)).normalize()
    return pd.to_datetime(value, errors="coerce")


def journals_frame(journals: list[dict], period_start, period_end, contacts_by_source_id: dict | None = None) -> pd.DataFrame:
    """One row per journal line in [period_start, period_end], in the same
    shape as xero_reports.parse_account_transactions. Xero reports a
    journal line's NetAmount debit-positive; the contra account is derived
    from the other lines of the same journal (largest first) since a
    journal has no 'Related account' field - flagged for review when there
    is more than one, exactly as the export's 'and N more' is."""
    start, end = pd.Timestamp(period_start), pd.Timestamp(period_end)
    contacts_by_source_id = contacts_by_source_id or {}
    rows = []
    for j in journals:
        date = _journal_date(j.get("JournalDate"))
        if pd.isna(date) or date < start or date > end:
            continue
        lines = j.get("JournalLines") or []
        source = j.get("SourceType", "")
        for line in lines:
            others = [o for o in lines if o is not line and o.get("AccountCode", o.get("AccountID")) != line.get("AccountCode", line.get("AccountID"))]
            others.sort(key=lambda o: abs(o.get("NetAmount", 0) or 0), reverse=True)
            distinct = []
            for o in others:
                key = o.get("AccountCode") or o.get("AccountName")
                if key not in [d[0] for d in distinct]:
                    distinct.append((key, o.get("AccountName", "")))
            net = float(line.get("NetAmount", 0) or 0)
            code = line.get("AccountCode") or line.get("AccountName", "")
            rows.append({
                "date": date, "account_code": str(code), "account_name": line.get("AccountName", ""),
                "reference": j.get("Reference", "") or "", "description": line.get("Description", "") or "",
                "contact": contacts_by_source_id.get(j.get("SourceID"), ""),
                "source_type": SOURCE_TYPE_LABELS.get(source, source),
                "debit": max(net, 0.0), "credit": max(-net, 0.0), "vat_amount": float(line.get("TaxAmount", 0) or 0),
                "contra_code": str(distinct[0][0]) if distinct and distinct[0][0] and distinct[0][0] != distinct[0][1] else "",
                "contra_name": distinct[0][1] if distinct else "", "contra_needs_review": len(distinct) > 1,
            })
    df = pd.DataFrame(rows, columns=[c for c in NOMINAL_COLUMNS if c != "net"])
    df["net"] = df["debit"] - df["credit"]
    return df.sort_values(["account_code", "date"], kind="stable").reset_index(drop=True)


# ---------- checking the job's period against what a year-to-date TB can mean ----------

def financial_year_warning(organisation_response: dict, period_start, period_end) -> str | None:
    orgs = organisation_response.get("Organisations") or []
    if not orgs:
        return None
    day, month = orgs[0].get("FinancialYearEndDay"), orgs[0].get("FinancialYearEndMonth")
    if not day or not month:
        return None
    start, end = pd.Timestamp(period_start), pd.Timestamp(period_end)
    year_end = pd.Timestamp(year=end.year, month=month, day=min(day, calendar.monthrange(end.year, month)[1]))
    ends_on_year_end = abs((end - year_end).days) <= 3
    is_full_year = 360 <= (end - start).days + 1 <= 371
    if not (ends_on_year_end and is_full_year):
        return (f"This organisation's financial year ends {day} {datetime(2000, month, 1):%B}, but this job covers "
                f"{pd.Timestamp(period_start):%d %b %Y} to {end:%d %b %Y}. Xero's Trial Balance is year-to-date for its own "
                f"financial year, so the pulled balances may not represent exactly this period - check before relying on them.")
    return None


# ---------- frames <-> bytes (stored as job files) ----------

def frame_to_bytes(df: pd.DataFrame) -> bytes:
    payload = json.loads(df.to_json(orient="split", date_format="iso", date_unit="ms"))
    return json.dumps({"columns": payload["columns"], "data": payload["data"]}).encode()


def frame_from_bytes(blob: bytes) -> pd.DataFrame:
    payload = json.loads(blob)
    df = pd.DataFrame(payload["data"], columns=payload["columns"])
    if "account_code" in df.columns:
        df["account_code"] = df["account_code"].astype(str)
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    return df
