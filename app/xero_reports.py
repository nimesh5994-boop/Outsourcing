"""Parsers for Xero's standard exported reports.

Xero's reports are not flat tables: they carry 3-4 title rows before the
real header, and the "Account Transactions" / "Aged Payables & Receivables
Detail" reports are grouped (a section header row per account/contact,
detail rows, then a "Total <name>" subtotal row). Generic column mapping
doesn't work on that shape, so these functions understand the Xero layout
directly. QBO/Sage/other exports still go through the generic mapping path
in parsers.py + mapping.py.
"""
import re

import pandas as pd

from app.parsers import DataSource, _to_numeric

TOTAL_PREFIX = "Total "


def _load_raw(source: DataSource) -> pd.DataFrame:
    raw = source.raw_dataframe()
    raw.columns = range(len(raw.columns))
    return raw


def _find_header_row(raw: pd.DataFrame, required_headers: list[str], max_scan: int = 15) -> int:
    for i in range(min(max_scan, len(raw))):
        row_vals = {str(v).strip().lower() for v in raw.iloc[i].tolist() if v is not None}
        if all(h.lower() in row_vals for h in required_headers):
            return i
    raise ValueError(f"Could not locate a header row containing {required_headers} in the first {max_scan} rows.")


def _sliced_with_header(raw: pd.DataFrame, header_row: int) -> pd.DataFrame:
    header = [str(v).strip() if v is not None else "" for v in raw.iloc[header_row].tolist()]
    body = raw.iloc[header_row + 1:].copy()
    body.columns = header
    return body.reset_index(drop=True)


_AS_AT = re.compile(r"as at\s+(.+)", re.IGNORECASE)
_FOR_PERIOD = re.compile(r"for the period\s+(.+?)\s+to\s+(.+)", re.IGNORECASE)


def extract_period_info(source: DataSource) -> dict:
    """Best-effort read of the period a Xero report actually covers, from its
    title rows (e.g. 'As at 31 December 2025' or 'For the period 1 January
    2025 to 31 December 2025') - so an uploaded file can be checked against
    the accounting period the job was set up for, rather than trusted blind."""
    raw = _load_raw(source)
    text_rows = [str(v).strip() for v in raw.iloc[:5, 0].tolist() if isinstance(v, str)]
    for text in text_rows:
        m = _FOR_PERIOD.search(text)
        if m:
            start, end = pd.to_datetime(m.group(1), errors="coerce", dayfirst=True), pd.to_datetime(m.group(2), errors="coerce", dayfirst=True)
            if pd.notna(start) and pd.notna(end):
                return {"kind": "range", "start": start.date(), "end": end.date(), "raw_text": text}
        m = _AS_AT.search(text)
        if m:
            as_of = pd.to_datetime(m.group(1), errors="coerce", dayfirst=True)
            if pd.notna(as_of):
                return {"kind": "as_of", "end": as_of.date(), "raw_text": text}
    return {"kind": "unknown", "raw_text": " | ".join(text_rows)}


def check_period(period_info: dict, expected_end, tolerance_days: int = 5) -> dict:
    """Compares an extracted report period against the job's declared period
    end date. expected_end may be None (nothing declared to check against)."""
    if period_info.get("kind") == "unknown":
        return {"status": "unknown", "message": "Could not detect a reporting period from this file - check it covers the right dates."}
    if expected_end is None:
        return {"status": "unknown", "message": f"Report covers: {period_info.get('raw_text', '')}"}

    found_end = period_info.get("end")
    if found_end is None:
        return {"status": "unknown", "message": f"Report covers: {period_info.get('raw_text', '')}"}

    delta = abs((found_end - expected_end).days)
    if delta <= tolerance_days:
        return {"status": "ok", "message": f"Report period matches ({period_info.get('raw_text', '')})."}
    return {
        "status": "mismatch",
        "message": f"Report says '{period_info.get('raw_text', '')}' but the job's period ends {expected_end:%d %b %Y} - check you uploaded the right file/period.",
    }


def parse_trial_balance(source: DataSource) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Returns (tb_current, tb_comparative). Xero's standard TB export embeds
    the comparative year as a single net (debit-positive/credit-negative)
    column, e.g. '31 Dec 2024'."""
    raw = _load_raw(source)
    header_row = _find_header_row(raw, ["Account Code", "Account", "Account Type"])
    body = _sliced_with_header(raw, header_row)

    known = {"Account Code", "Account", "Account Type", "Debit - Year to date", "Credit - Year to date"}
    comparative_col = next((c for c in body.columns if c not in known and c != ""), None)

    # Some accounts (typically directly-connected bank feeds) have no nominal
    # code assigned in Xero and are identified by name only - keep those rows,
    # just drop the blank separator row and the report's own "Total" row.
    code_str = body["Account Code"].astype(str).str.strip()
    is_total_row = code_str.str.lower() == "total"
    is_blank_row = body["Account Code"].isna() & body["Account"].isna()
    body = body[~is_total_row & ~is_blank_row]

    account_name = body["Account"].fillna("").astype(str).str.strip()
    account_code = body["Account Code"].fillna("").astype(str).str.strip()
    # directly-connected bank accounts often have no nominal code in Xero -
    # fall back to the account name so every row keeps a usable, non-blank key
    account_code = account_code.where(account_code != "", account_name)

    out = pd.DataFrame({
        "account_code": account_code,
        "account_name": account_name,
        "account_type": body["Account Type"].astype(str).str.strip(),
        "debit": _to_numeric(body["Debit - Year to date"]),
        "credit": _to_numeric(body["Credit - Year to date"]),
    })
    out["balance"] = out["debit"] - out["credit"]
    tb_current = out.reset_index(drop=True)

    if comparative_col is not None:
        comp_balance = _to_numeric(body[comparative_col])
        comparative = pd.DataFrame({
            "account_code": out["account_code"],
            "account_name": out["account_name"],
            "account_type": out["account_type"],
            "debit": comp_balance.clip(lower=0),
            "credit": (-comp_balance).clip(lower=0),
        })
        comparative["balance"] = comparative["debit"] - comparative["credit"]
        tb_comparative = comparative.reset_index(drop=True)
    else:
        tb_comparative = pd.DataFrame(columns=tb_current.columns)

    return tb_current, tb_comparative


_SECTION_CODE_SUFFIX = re.compile(r"\s*\(\d+\)\s*$")
_RELATED_ACCOUNT_CODE = re.compile(r"^\s*(\d+)\s*-\s*(.+)$")
_AND_MORE = re.compile(r"and \d+ more$", re.IGNORECASE)


def parse_account_transactions(source: DataSource) -> pd.DataFrame:
    """Xero 'Account Transactions' report: grouped by account, one section
    per nominal code. Each detail row already carries its own Account Code
    and Account Type, but NOT an account name column - that only appears in
    the section header text, so it's forward-filled from there."""
    raw = _load_raw(source)
    header_row = _find_header_row(raw, ["Date", "Source", "Account Code"])
    body = _sliced_with_header(raw, header_row)

    dates = pd.to_datetime(body["Date"], errors="coerce")
    is_detail = dates.notna()

    # The report is grouped into one section per account: a header row (name,
    # sometimes with a code suffix like "(2520)"), its detail rows, then a
    # "Total <name>" row. Detail rows carry their own Account Code *except*
    # for allocation-type entries (payments/receipts against an invoice),
    # which are blank - so the code is taken as the most common non-blank
    # value seen anywhere in that section, not read row-by-row.
    account_names = pd.Series(index=body.index, dtype=object)
    section_ids = pd.Series(index=body.index, dtype="Int64")
    current_name, section_id = None, -1
    first_col = body.iloc[:, 0]
    for idx in body.index:
        if is_detail[idx]:
            account_names[idx] = current_name
            section_ids[idx] = section_id
        elif isinstance(first_col[idx], str) and not first_col[idx].startswith(TOTAL_PREFIX):
            current_name = _SECTION_CODE_SUFFIX.sub("", first_col[idx]).strip()
            section_id += 1

    detail = body[is_detail].copy()
    detail["__account_name__"] = account_names[is_detail]
    detail["__section_id__"] = section_ids[is_detail]

    row_code = detail["Account Code"].astype(str).str.strip()
    row_code = row_code.where(row_code.notna() & (row_code != "") & (row_code.str.lower() != "nan"), None)
    section_code = row_code.groupby(detail["__section_id__"]).transform(
        lambda s: s.mode().iat[0] if not s.mode().empty else None
    )
    detail["__account_code__"] = section_code.fillna(detail["__account_name__"])

    related = detail.get("Related account", pd.Series("", index=detail.index)).fillna("").astype(str)

    def split_related(val: str) -> tuple[str, str, bool]:
        val = val.strip()
        if not val:
            return "", "", False
        needs_review = bool(_AND_MORE.search(val))
        first_item = val.split(",")[0].strip()
        m = _RELATED_ACCOUNT_CODE.match(first_item)
        if m:
            return m.group(1), m.group(2).strip(), needs_review
        return "", first_item, needs_review

    contra = related.apply(split_related)
    out = pd.DataFrame({
        "date": dates[is_detail].values,
        "account_code": detail["__account_code__"],
        "account_name": detail["__account_name__"],
        "reference": detail.get("Reference", detail.get("Invoice Number", "")).fillna(""),
        "description": detail.get("Description", "").fillna(""),
        "contact": detail.get("Contact", "").fillna(""),
        "source_type": detail.get("Source", "").fillna(""),
        "debit": _to_numeric(detail.get("Debit", 0)),
        "credit": _to_numeric(detail.get("Credit", 0)),
        "contra_code": [c[0] for c in contra],
        "contra_name": [c[1] for c in contra],
        "contra_needs_review": [c[2] for c in contra],
    })
    out["net"] = out["debit"] - out["credit"]
    return out.reset_index(drop=True)


_BUCKET_COLUMNS = [
    ("current", "Current"), ("bucket_1", "< 1 Month"), ("bucket_2", "1 Month"),
    ("bucket_3", "2 Months"), ("bucket_4", "3 Months"), ("older", "Older"), ("total", "Total"),
]


def parse_aged_report(source: DataSource, party_field: str) -> pd.DataFrame:
    """Xero 'Aged Payables/Receivables Detail' report: grouped by contact,
    invoice-level rows, then a 'Total <name>' subtotal row per contact and a
    final grand 'Total' + 'Percentage of total' row. The per-contact subtotal
    rows use un-evaluated formulas (e.g. '=E9') in Xero's own export, which
    read back as blank/zero - so contact totals are summed from the detail
    rows directly rather than trusted from the subtotal row."""
    raw = _load_raw(source)
    header_row = _find_header_row(raw, ["Current", "Older", "Total"])
    body = _sliced_with_header(raw, header_row)

    dates = pd.to_datetime(body["Invoice Date"], errors="coerce")
    is_detail = dates.notna()

    first_col = body.iloc[:, 0]
    is_skip = first_col.astype(str).str.startswith(TOTAL_PREFIX) | first_col.astype(str).str.startswith("Percentage")

    party = pd.Series(index=body.index, dtype=object)
    current_party = None
    for idx in body.index:
        if is_detail[idx]:
            party[idx] = current_party
        elif not is_skip[idx] and isinstance(first_col[idx], str) and first_col[idx].strip():
            current_party = first_col[idx].strip()

    detail = body[is_detail].copy()
    detail[party_field] = party[is_detail]
    for canonical, source_col in _BUCKET_COLUMNS:
        detail[canonical] = _to_numeric(detail[source_col]) if source_col in detail.columns else 0.0

    out = detail.groupby(party_field, as_index=False)[[c for c, _ in _BUCKET_COLUMNS]].sum()
    return out.reset_index(drop=True)


def derive_pl_bs_from_tb(tb: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Xero's TB carries an Account Type per row (Sales/Direct Costs/Overhead/
    Expense = P&L; Bank/Current Asset/Fixed Asset/Current Liability/
    Liability/Equity = B/S), so P&L and B/S can be derived directly instead
    of requiring separate uploads."""
    if tb is None or tb.empty:
        return pd.DataFrame(), pd.DataFrame()
    pl_types = {"sales", "direct costs", "overhead", "overheads", "expense", "income", "other income"}
    df = tb.copy()
    df["_type_l"] = df["account_type"].str.lower()
    is_pl = df["_type_l"].isin(pl_types)

    pl = df[is_pl].copy()
    pl["amount"] = pl["credit"] - pl["debit"]  # income positive, expenses negative
    pl["category"] = pl["account_type"]
    pl = pl[["account_code", "account_name", "category", "amount"]].reset_index(drop=True)

    bs = df[~is_pl].copy()
    bs["amount"] = bs["debit"] - bs["credit"]  # assets positive, liabilities/equity negative
    bs["category"] = bs["account_type"]
    bs = bs[["account_code", "account_name", "category", "amount"]].reset_index(drop=True)
    return pl, bs
