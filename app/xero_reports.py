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
from datetime import datetime

import pandas as pd

from app.parsers import DataSource, _to_numeric

TOTAL_PREFIX = "Total "


def _load_raw(source: DataSource) -> pd.DataFrame:
    # .copy() matters: DataSource caches and returns the same DataFrame
    # object on every call, so renaming columns in place here would
    # permanently corrupt it for every later raw_columns()/raw_dataframe()
    # call on this same source - including the generic column-mapping
    # fallback that runs right after a failed Xero-native parse attempt
    # (e.g. platform was guessed/selected as Xero but the file wasn't a
    # genuine native export) - that fallback would then see integer
    # column positions instead of real headers and map nothing at all.
    raw = source.raw_dataframe().copy()
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
# Most Xero reports (TB, P&L) write "For the period 1 January 2025 to 31
# December 2025", but the VAT Return export writes the exact same thing as
# "For the period 01 Feb 2025 - 28 Feb 2025" - a dash, not the word "to".
# Found live: this meant extract_period_info returned "unknown" for every
# single real VAT Return file tested (six full periods, zero exceptions),
# which silently broke guess_period's date-based bucketing for VAT returns
# entirely - with no date signal, it fell back to the order-dependent
# "second upload of this type = comparative" heuristic, so period bucketing
# depended on upload order rather than the file's own dates. Matching
# either separator fixes both the label text and the underlying bug.
_FOR_PERIOD = re.compile(r"for the period\s+(.+?)\s+(?:to|-|–|—)\s+(.+)", re.IGNORECASE)


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
        "vat_amount": _to_numeric(detail.get("VAT", 0)),
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


_AGED_TITLE_KEYWORDS = {
    "customer": ("receivable", "debtor"),
    "supplier": ("payable", "creditor"),
}


def classify_aged_report_party(text: str) -> str | None:
    """Reads free text (a report title, or any text pulled from near the
    top of a file) and returns 'customer' (a Receivables/Debtors report)
    or 'supplier' (a Payables/Creditors report) - or None if the text
    doesn't clearly say either. Shared between parse_aged_report's own
    Xero-native title check below and the generic (non-Xero-native, e.g.
    a PDF export) upload path in main.py, which hits the exact same
    ambiguity for the exact same reason: an Aged Payables Detail export
    and an Aged Receivables Detail export are structurally identical past
    their own title line - same grouped-by-contact shape, same bucket
    columns - so this text is the only reliable signal either path has to
    tell them apart. Found live via the generic path specifically: a real
    Aged Payables Detail PDF, once table-extracted, has no column header
    text that says "payable" or "creditor" anywhere (the bucket columns
    are just Current/1 Month/.../Total, identical to a receivables
    export), so the ordinary alias-scoring classifier had nothing to
    prefer aged_creditors over aged_debtors with and picked the wrong one."""
    lowered = str(text).strip().lower()
    for party_field, keywords in _AGED_TITLE_KEYWORDS.items():
        if any(k in lowered for k in keywords):
            return party_field
    return None


def parse_aged_report(source: DataSource, party_field: str) -> pd.DataFrame:
    """Xero 'Aged Payables/Receivables Detail' report: grouped by contact,
    invoice-level rows, then a 'Total <name>' subtotal row per contact and a
    final grand 'Total' + 'Percentage of total' row. The per-contact subtotal
    rows use un-evaluated formulas (e.g. '=E9') in Xero's own export, which
    read back as blank/zero - so contact totals are summed from the detail
    rows directly rather than trusted from the subtotal row.

    An Aged Payables Detail export and an Aged Receivables Detail export are
    structurally identical past the title row - same grouped-by-contact
    shape, same Current/1 Month/.../Total bucket columns - so nothing below
    this point can tell them apart. Without checking the title, this parser
    happily parses EITHER file as EITHER party_field, and try_xero_native's
    "first Xero-native parser that doesn't raise wins" logic then picks
    whichever of aged_debtors/aged_creditors it happens to try first - found
    live: a real client's Aged Receivables Detail export got auto-detected
    as aged_creditors (mixing debtor invoices into the creditors control
    account check, and leaving debtors reconciliation showing "no aged
    report uploaded" despite one being confirmed). The title row ('Aged
    Payables Detail' / 'Aged Receivables Detail') is the one signal that
    actually distinguishes them, so it's checked first and rejected loudly
    (not silently parsed as the wrong type) when it doesn't match. That
    title text is read via source.raw_columns()[0] rather than a row of
    _load_raw()'s DataFrame: pandas treats the file's very first row as the
    column header, so the title ends up as the raw column name (source.
    raw_columns() returns it as 'Aged Payables Detail', 'Unnamed: 1', ...)
    and is gone from row data entirely by the time _load_raw() re-indexes
    columns to plain integers."""
    columns = source.raw_columns()
    title = str(columns[0]).strip().lower() if columns and columns[0] is not None else ""
    keywords = _AGED_TITLE_KEYWORDS[party_field]
    if title and not any(k in title for k in keywords):
        raise ValueError(f"Title row {title!r} doesn't match an Aged {'Receivables' if party_field == 'customer' else 'Payables'} Detail report.")

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


_VAT_BOX_NUMBERS = {str(i) for i in range(1, 10)}
_VAT_REQUIRED_BOXES = {"box1", "box3", "box4", "box5"}


def parse_vat_return_box_summary(source: DataSource) -> pd.DataFrame:
    """Xero's exported 'VAT Return' report - not a table at all, but a
    vertical label/box-number/value listing (one row per HMRC box, e.g.
    ['VAT due in the period on sales and other outputs', '1', 492.36]),
    with title/client-name/period/scheme-detail rows above and below it
    that carry no box number. Found live: a real client's VAT Return
    export dropped into the VAT Return upload section came through as a
    3-sheet workbook (this report plus two structurally unrelated detail
    sheets - see main.py's XERO_NATIVE_REPORT_TYPES handling), and the
    generic column-mapper found nothing to map on ANY of the three (no
    row in this shape looks like a header row with per-box columns), so
    the top-level VAT cross-check silently had nothing to work with.
    Returns a single-row DataFrame with columns box1..box9 (0.0 for any
    box this export happens to omit, e.g. an export from a business with
    no EU trade never mentions boxes 8/9)."""
    raw = _load_raw(source)
    if raw.shape[1] < 3:
        raise ValueError("Not a VAT Return box summary export - expected at least 3 columns (label, box number, value).")

    boxes: dict[str, object] = {}
    for _, row in raw.iterrows():
        box_no = str(row.iloc[1]).strip() if row.iloc[1] is not None else ""
        if box_no in _VAT_BOX_NUMBERS:
            boxes[f"box{box_no}"] = row.iloc[2]

    if not _VAT_REQUIRED_BOXES.issubset(boxes.keys()):
        raise ValueError("Not a VAT Return box summary export - couldn't find boxes 1, 3, 4 and 5.")

    for i in range(1, 10):
        boxes.setdefault(f"box{i}", 0.0)
    out = pd.DataFrame([boxes])
    for col in out.columns:
        out[col] = _to_numeric(out[col])
    return out


_VAT_BOX_ANY_CELL = re.compile(r"^Box\s+(\d)$", re.IGNORECASE)
_VAT_BOX_DETAIL_COLUMNS = ("date", "reference", "contact", "description", "net_amount", "vat_amount")


def parse_vat_box_transactions(source: DataSource) -> dict[int, pd.DataFrame]:
    """Xero's 'Transactions by VAT Box' export: the actual postings behind
    each HMRC box on the VAT Return - the same file parse_vat_return_box_
    summary reads the box TOTALS from, but this sheet has the transaction-
    level detail behind boxes 1 and 4 specifically (sales/purchases with
    VAT), which is exactly the vat_filed_sales/vat_filed_purchases data
    the VAT Reconciliation workspace (app/vat_reconciliation.py) needs -
    normally sourced from a separately-exported filed-return detail file,
    but a client's Xero VAT Return export already has it, box-tagged, in
    one sheet.

    Structure: a 'Box N - <description>' row (with that box's total, no
    date), then one or more tax-rate sub-groups (e.g. '20% (VAT on
    Income)', 'Zero Rated Expenses') each with their own repeated 'Date |
    Account | Reference | Details | VAT | Net' header and detail rows.
    Deliberately doesn't trust a fixed column position for any of this -
    found live across several real periods for the same client that a
    longer (quarterly) period's export inserts an extra leading column
    holding the *current* box number, repeated on every single row of
    that box's section (not just its header) - shifting Date/Account/../
    Net one column to the right of where a shorter (monthly) period's
    export puts them, and turning a fixed-column "is this cell exactly
    'Box N'" check into a false match on every ordinary data row, since
    that leading column keeps saying 'Box 1' for 160+ rows straight. So:
    1) every header row is found by CONTENT ('Date' immediately followed
    by 'Account', wherever that pair of cells actually sits) rather than
    assumed to be at column 0/1, and each one's own detail rows are read
    relative to ITS OWN column offset; 2) which box a header belongs to
    is resolved by scanning backwards from it for the nearest row with a
    'Box N' cell ANYWHERE in it - correct whether that cell appears once
    (monthly) or is repeated down every row of the section (quarterly),
    since either way the nearest one behind a given header is that
    header's own box. Boxes 6/7 repeat the exact same underlying
    transactions as boxes 1/4 (net-only, no VAT column, since 6/7 are
    "excluding VAT" totals) - skipped as pure duplicates of 1/4's own Net
    column. 'Details' is used for both contact and description: on the
    sales side (Box 1) it's reliably a customer/payee name, but on the
    purchases side (Box 4) real exports mix genuine vendor names with
    free-text card-transaction descriptions ("BP S/S - BP", "Starbucks -
    Subsistence") - an honest limitation of the source data, not
    something parsing can fix, and the reason Box 4 GL-matching may find
    more "no match" candidates than Box 1 does on the same file.

    Returns {box_number: DataFrame} in vat_filed_sales/vat_filed_purchases'
    own canonical shape (date, reference, contact, description,
    net_amount, vat_amount) - only for whichever of boxes 1/4 this file
    actually has transaction detail for (an export can have one without
    the other, e.g. a period with only sales or only purchases)."""
    raw = _load_raw(source)
    n_rows, n_cols = raw.shape
    if n_cols < 5:
        raise ValueError("Not a 'Transactions by VAT Box' export - expected at least 5 columns.")

    def cell(r: int, c: int):
        return raw.iat[r, c] if 0 <= c < n_cols else None

    def cell_str(r: int, c: int) -> str:
        v = cell(r, c)
        return str(v).strip() if v is not None else ""

    header_positions: list[tuple[int, int]] = []  # (row, column offset of "Date")
    for r in range(n_rows):
        for c in range(n_cols - 1):
            if cell_str(r, c).lower() == "date" and cell_str(r, c + 1).lower() == "account":
                header_positions.append((r, c))
                break

    if not header_positions:
        raise ValueError("Not a 'Transactions by VAT Box' export - no 'Date | Account | ...' header row found.")

    box_marks: list[tuple[int, int]] = []  # (row, box number), every "Box N" cell anywhere
    for r in range(n_rows):
        for c in range(n_cols):
            m = _VAT_BOX_ANY_CELL.match(cell_str(r, c))
            if m:
                box_marks.append((r, int(m.group(1))))
                break

    def box_for_header(header_row: int) -> int | None:
        for r, box in reversed(box_marks):
            if r <= header_row:
                return box
        return None

    rows_by_box: dict[int, list[dict]] = {}
    for header_row, offset in header_positions:
        box = box_for_header(header_row)
        if box not in (1, 4):  # boxes 6/7 duplicate 1/4's own transactions net-only - skip
            continue
        # A Northern Ireland/EU acquisition affects boxes 2 and 4 (and 9)
        # from the SAME underlying transaction, so Xero cross-references it
        # into the Box 4 section as a duplicated pair of tax-rate sub-
        # groups: "EC Acquisitions (NN%)" (the acquisition-due side - Box
        # 2's own transaction, not a Box 4 addition) immediately followed
        # by "EC Acquisitions (NN%) Reclaimed VAT" (the actual Box 4
        # input-VAT reclaim) - found live with BOTH showing the identical
        # VAT figure for the identical transaction, which double-counted
        # it into Box 4's total when both were collected. Only the
        # "Reclaimed VAT" one belongs to Box 4; skip the acquisition-due
        # one (its own real home, Box 2, isn't collected here at all).
        sub_group_label = cell_str(header_row - 1, offset).lower()
        if box == 4 and "ec acquisitions" in sub_group_label and "reclaim" not in sub_group_label:
            continue
        r = header_row + 1
        while r < n_rows:
            date_val = cell(r, offset)
            # already a real datetime for a genuine date cell (openpyxl
            # reads Excel dates that way) - stringifying it first before
            # re-parsing needlessly ambiguates an unambiguous value and
            # trips pandas' dayfirst-format warning for no reason.
            if isinstance(date_val, (pd.Timestamp, datetime)):
                date = pd.Timestamp(date_val)
            else:
                date = pd.to_datetime(cell_str(r, offset), dayfirst=True, errors="coerce")
            if pd.isna(date):
                break  # blank separator row, the next sub-group's tax-rate label, or another header
            details = cell(r, offset + 3)
            rows_by_box.setdefault(box, []).append({
                "date": date, "reference": cell(r, offset + 2) if cell(r, offset + 2) is not None else "",
                "contact": details if details is not None else "", "description": details if details is not None else "",
                "net_amount": cell(r, offset + 5), "vat_amount": cell(r, offset + 4),
            })
            r += 1

    if not rows_by_box:
        raise ValueError("Not a 'Transactions by VAT Box' export - no Box 1/Box 4 transaction detail found.")

    out = {}
    for box, box_rows in rows_by_box.items():
        df = pd.DataFrame(box_rows, columns=list(_VAT_BOX_DETAIL_COLUMNS))
        df["net_amount"] = _to_numeric(df["net_amount"])
        df["vat_amount"] = _to_numeric(df["vat_amount"])
        out[box] = df.reset_index(drop=True)
    return out


# Xero's TB carries an Account Type per row (Revenue/Sales/Direct Costs/
# Overhead/Expense = P&L; Bank/Current Asset/Fixed Asset/Current
# Liability/Liability/Equity = B/S). "Revenue" is Xero's own standard
# Account Type label for sales accounts (distinct from the account
# *name*, which is often "Sales") - a real bug found live: without it
# here, those accounts silently fell through to the B/S side instead,
# understating Turnover to zero and throwing off the B/S balance check
# by the same amount. Exported (not a local variable) so every place
# that needs the same P&L/B/S split - derive_pl_bs_from_tb below, and
# tb_tieout.py's opening-balance logic (a P&L account has no carried-
# forward opening balance the way a balance-sheet one does) - shares
# one definition instead of risking the two silently drifting apart.
PL_ACCOUNT_TYPES = {"sales", "revenue", "direct costs", "overhead", "overheads", "expense", "income", "other income"}


def derive_pl_bs_from_tb(tb: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """See PL_ACCOUNT_TYPES above for which Account Types are treated as
    P&L vs Balance Sheet."""
    if tb is None or tb.empty:
        return pd.DataFrame(), pd.DataFrame()
    df = tb.copy()
    df["_type_l"] = df["account_type"].str.lower()
    is_pl = df["_type_l"].isin(PL_ACCOUNT_TYPES)

    pl = df[is_pl].copy()
    pl["amount"] = pl["credit"] - pl["debit"]  # income positive, expenses negative
    pl["category"] = pl["account_type"]
    pl = pl[["account_code", "account_name", "category", "amount"]].reset_index(drop=True)

    bs = df[~is_pl].copy()
    bs["amount"] = bs["debit"] - bs["credit"]  # assets positive, liabilities/equity negative
    bs["category"] = bs["account_type"]
    bs = bs[["account_code", "account_name", "category", "amount"]].reset_index(drop=True)
    return pl, bs
