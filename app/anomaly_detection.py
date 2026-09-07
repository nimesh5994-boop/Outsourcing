"""Cross-transaction anomaly detection over nominal activity.

recon.nominal_activity_review flags a transaction based on what's on that
one row (suspense code, round-sum journal, a "to be reviewed" narrative).
The checks here look across *many* transactions for the same contact to
find patterns a single row can't reveal on its own - the canonical example
being a supplier whose spend is split across nominal codes in a way that
looks like miscoding rather than genuinely mixed spend: "BT: 10 postings to
Telephone, 2 to Light & Heat" is a strong signal the 2 are wrong, not that
BT provides two different services.

Every check returns a recon.ReconResult (status/message/detail), the same
shape every other check in this system uses, so these plug straight into
the existing results list, index sheet, and recon-sheet builder with no
new plumbing.
"""
import pandas as pd

from app.fixed_assets import _MIGRATION_KEYWORDS
from app.recon import ReconResult

MIN_TRANSACTIONS_FOR_PATTERN = 5  # need a decent sample before trusting "this contact codes consistently"
DOMINANT_CODE_MIN_SHARE = 0.75    # one code must account for at least this share of a contact's postings
MINORITY_CODE_MAX_SHARE = 0.25    # ...for a different code used this rarely to look like an outlier, not mixed spend


def _amount(df: pd.DataFrame) -> pd.Series:
    return (df["debit"].fillna(0) - df["credit"].fillna(0)).round(2)


def _looks_like_migration_entry(row: pd.Series) -> bool:
    """Same signal fixed_assets.py already uses to tell a genuine new
    purchase from a historic-balance migration journal (see its own
    _MIGRATION_KEYWORDS comment): a suspense/clearing line created when a
    client's opening balances were imported is nearly always labelled as
    one somewhere - its own account name, description or reference. Reused
    here for the same reason it matters for duplicates: Xero's own import
    mechanism posts a mirrored pair (one line to the real control account,
    one to a suspense "Opening Balance" account) for every migrated
    balance, sharing the same date/contact/amount/reference by design -
    found live on a real client's file, where this pattern alone accounted
    for a large share of a 3489-transaction "duplicate" flag despite being
    a normal, expected artifact of the import, not a risk."""
    text = f"{row.get('account_name', '')} {row.get('description', '')} {row.get('reference', '')}".lower()
    return any(k in text for k in _MIGRATION_KEYWORDS)


def _has_text(series: pd.Series) -> pd.Series:
    return series.notna() & (series.astype(str).str.strip() != "")


def contact_coding_consistency(nominal_activity: pd.DataFrame) -> ReconResult:
    """For each contact with enough transaction volume, finds a nominal
    code used for a small minority of their postings while a different
    code clearly dominates, and flags those minority transactions with the
    dominant code as the likely correct allocation - the BT-posted-to-
    Telephone-and-Light&Heat pattern."""
    name = "Contact coding consistency (possible miscoding)"
    if nominal_activity is None or nominal_activity.empty:
        return ReconResult(name, "n/a", "No nominal activity uploaded.")
    if "contact" not in nominal_activity.columns:
        return ReconResult(name, "n/a", "No contact/payee detail available in the nominal activity uploaded.")

    df = nominal_activity[_has_text(nominal_activity["contact"])].copy()
    if df.empty:
        return ReconResult(name, "n/a", "No transactions have a contact/payee recorded.")
    df["amount"] = _amount(df)

    flagged_rows = []
    for contact, g in df.groupby("contact"):
        if len(g) < MIN_TRANSACTIONS_FOR_PATTERN:
            continue
        code_counts = g.groupby(["account_code", "account_name"]).size().sort_values(ascending=False)
        if len(code_counts) < 2:
            continue  # only ever posted to one code - nothing to compare against

        total = int(code_counts.sum())
        dominant_code, dominant_name = code_counts.index[0]
        dominant_count = int(code_counts.iloc[0])
        if (dominant_count / total) < DOMINANT_CODE_MIN_SHARE:
            continue  # no single code clearly dominates - looks like genuinely mixed spend, not a coding slip

        minority = code_counts.iloc[1:]
        minority = minority[(minority / total) <= MINORITY_CODE_MAX_SHARE]
        if minority.empty:
            continue

        for (code, acct_name), count in minority.items():
            txns = g[(g["account_code"].astype(str) == str(code)) & (g["account_name"] == acct_name)]
            for _, t in txns.iterrows():
                flagged_rows.append({
                    "Contact": contact,
                    "Date": t.get("date"),
                    "Reference": t.get("reference"),
                    "Description": t.get("description"),
                    "Amount": t["amount"],
                    "Posted to": f"{code} - {acct_name}",
                    "This code's share of contact's postings": f"{int(count)} of {total}",
                    "Likely correct code": f"{dominant_code} - {dominant_name}",
                    "Dominant code's share of contact's postings": f"{dominant_count} of {total}",
                })

    detail = pd.DataFrame(flagged_rows)
    status = "ok" if detail.empty else "review"
    msg = (
        "No contact shows a lopsided coding pattern - either everyone codes consistently, "
        "or spread across codes looks like genuinely mixed spend rather than a slip."
        if detail.empty else
        f"{len(detail)} transaction(s) look miscoded: posted to a nominal code used only rarely for that "
        f"contact, when a different code clearly dominates their history. Review and reallocate to the "
        f"suggested code if it's a genuine slip."
    )
    return ReconResult(name, status, msg, detail)


def duplicate_transactions(nominal_activity: pd.DataFrame) -> ReconResult:
    """Flags likely duplicate postings: the same contact billed for the
    same date/amount(/VAT) more than once, or the same reference/invoice
    number used more than once for the same contact on the same date -
    either can be entirely legitimate (a genuine repeat charge, a credit
    note matching an invoice), so this is a "check before you trust it"
    flag, not a definite error.

    Keyed on date + contact + amount + VAT (when the upload carries a VAT
    amount - not every export does) rather than just contact+amount: found
    live against a real client whose reference numbers are short, reused
    sequence numbers (order/pallet numbers, not unique invoice IDs) - the
    same reference recurred for the same contact across entirely different
    invoices weeks apart, and without date narrowing the key, this alone
    produced 3489 flagged "duplicates" on one real file, the overwhelming
    majority not duplicates at all. Adding VAT narrows the amount-based key
    further: two coincidentally same-day, same-amount postings to the same
    contact with different VAT treatment are almost certainly two different
    transactions, not one entered twice.

    Also flags (via the Advisory column, not by excluding the row - see
    _looks_like_migration_entry) the other major source of false positives
    found on the same file: Xero's own opening-balance import posts a
    mirrored pair of lines (the real control account + a suspense "Opening
    Balance" account) for every migrated historic balance, sharing date/
    contact/amount/reference by design - a known, expected artifact of the
    import, not a risk, so a preparer shouldn't need to individually
    re-investigate every one of these to conclude that."""
    name = "Duplicate transaction check"
    if nominal_activity is None or nominal_activity.empty:
        return ReconResult(name, "n/a", "No nominal activity uploaded.")
    if "contact" not in nominal_activity.columns:
        return ReconResult(name, "n/a", "No contact/payee detail available in the nominal activity uploaded.")

    df = nominal_activity.copy()
    df["amount"] = _amount(df)
    has_contact = _has_text(df["contact"])
    # A generic (non-native) mapping defaults an unmapped vat_amount column
    # to all-zero rather than leaving it absent, so column presence alone
    # can't tell "this file genuinely has no VAT data" from "it's just
    # zero" - any non-zero VAT anywhere is the signal the file actually
    # carries it.
    has_vat = "vat_amount" in df.columns and df["vat_amount"].fillna(0).abs().sum() > 0.005
    if has_vat:
        df["_vat"] = df["vat_amount"].fillna(0).round(2)

    dup_amount_mask = pd.Series(False, index=df.index)
    dup_ref_mask = pd.Series(False, index=df.index)
    if "date" in df.columns:
        amount_key_cols = ["contact", "date", "amount"] + (["_vat"] if has_vat else [])
        amount_key = has_contact & (df["amount"] != 0)
        dup_amount_mask = df.duplicated(subset=amount_key_cols, keep=False) & amount_key
    if "reference" in df.columns and "date" in df.columns:
        # same reference + same nominal code + same (signed) amount + same
        # date: a single invoice/bill legitimately posts the same
        # reference number to more than one code (e.g. Sales credit +
        # Debtors debit) and an invoice-then-its-payment share a reference
        # on the same code with opposite signs - neither is a duplicate, so
        # code and signed amount matching too is what actually narrows this
        # down to "the same line looks like it was entered twice." Date is
        # included for the reason in this function's docstring: a short,
        # reused reference number is a poor duplicate signal across
        # different dates on its own.
        ref_key_cols = ["contact", "reference", "account_code", "amount", "date"] + (["_vat"] if has_vat else [])
        ref_key = has_contact & _has_text(df["reference"]) & (df["amount"] != 0)
        dup_ref_mask = df.duplicated(subset=ref_key_cols, keep=False) & ref_key

    combined = dup_amount_mask | dup_ref_mask
    if not combined.any():
        vat_note = "/VAT" if has_vat else ""
        return ReconResult(name, "ok", f"No same contact/date/amount{vat_note} or same-date repeated reference numbers found.")

    flagged = df[combined].copy()
    reasons = []
    for idx in flagged.index:
        parts = []
        if dup_amount_mask.get(idx, False):
            parts.append(f"same contact, date and amount{' and VAT' if has_vat else ''} posted more than once")
        if dup_ref_mask.get(idx, False):
            parts.append("same reference/invoice number used more than once for this contact on the same date")
        reason = "; ".join(parts)
        reasons.append(reason[:1].upper() + reason[1:])  # capitalize only the first letter - .capitalize() would also lowercase "VAT"
    flagged["Flag reason"] = reasons
    flagged["Advisory"] = [
        "Looks like a known opening-balance migration entry - often safe to disregard" if _looks_like_migration_entry(row) else ""
        for _, row in flagged.iterrows()
    ]
    flagged["Reviewed - genuine repeat? (to complete)"] = ""

    sort_cols = [c for c in ("contact", "date") if c in flagged.columns]
    if sort_cols:
        flagged = flagged.sort_values(sort_cols)

    cols = [c for c in ("date", "account_code", "account_name", "reference", "description", "contact") if c in flagged.columns]
    cols += ["amount"] + (["_vat"] if has_vat else []) + ["Flag reason", "Advisory", "Reviewed - genuine repeat? (to complete)"]
    detail = flagged[cols].rename(columns={
        "date": "Date", "account_code": "Nominal Code", "account_name": "Account Name",
        "reference": "Reference", "description": "Description", "contact": "Contact", "amount": "Amount",
        "_vat": "VAT",
    })
    vat_note = "+VAT" if has_vat else ""
    msg = (
        f"{len(detail)} transaction(s) share a contact+date+amount{vat_note} or a repeated reference number "
        f"(same date) with another posting - review each: could be a genuine duplicate posting to correct, or "
        f"a legitimate repeat charge/credit note match."
    )
    migration_count = int((flagged["Advisory"] != "").sum())
    if migration_count:
        msg += (
            f" {migration_count} of these look like opening-balance migration entries (see Advisory column) - "
            f"these typically don't need individual review."
        )
    return ReconResult(name, "review", msg, detail)


def unusual_journal_posting_dates(nominal_activity: pd.DataFrame) -> ReconResult:
    """Flags manual journals posted on a weekend - restricted to journal-
    sourced entries specifically (not bank feeds or trading transactions,
    which legitimately happen any day of the week) to keep this a
    meaningful sense-check rather than noise for a business that trades
    weekends."""
    name = "Unusual posting date check (weekend manual journals)"
    if nominal_activity is None or nominal_activity.empty:
        return ReconResult(name, "n/a", "No nominal activity uploaded.")
    if "source_type" not in nominal_activity.columns or "date" not in nominal_activity.columns:
        return ReconResult(name, "n/a", "No source/date detail available to check posting dates.")

    df = nominal_activity.copy()
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    is_journal = df["source_type"].astype(str).str.lower().str.contains("journal", na=False)
    df = df[is_journal & df["date"].notna()]
    if df.empty:
        return ReconResult(name, "n/a", "No manual journal entries with a usable date found.")

    weekend = df[df["date"].dt.dayofweek >= 5].copy()
    if weekend.empty:
        return ReconResult(name, "ok", "No manual journals posted on a weekend date.")

    weekend["amount"] = _amount(weekend)
    cols = [c for c in ("date", "account_code", "account_name", "reference", "description", "contact") if c in weekend.columns]
    cols.append("amount")
    detail = weekend[cols].sort_values("date").rename(columns={
        "date": "Date", "account_code": "Nominal Code", "account_name": "Account Name",
        "reference": "Reference", "description": "Description", "contact": "Contact", "amount": "Amount",
    })
    msg = (
        f"{len(detail)} manual journal(s) posted on a weekend date - not necessarily wrong (a backdated or "
        f"scheduled entry can land on any day), but worth a quick sense-check that the date is genuine."
    )
    return ReconResult(name, "review", msg, detail)


def run_all_anomaly_checks(nominal_activity: pd.DataFrame) -> list[ReconResult]:
    return [
        contact_coding_consistency(nominal_activity),
        duplicate_transactions(nominal_activity),
        unusual_journal_posting_dates(nominal_activity),
    ]
