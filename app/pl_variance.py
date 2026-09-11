"""P&L variance review: current vs comparative movement on every P&L
account, with three pieces of extra context a bare TB-wide variance check
(recon.variance_analysis) doesn't give a reviewer:

  1. Which contacts/transactions are actually driving a flagged account's
     movement, and whether any of them also post to OTHER nominal codes -
     the "posted to multiple codes" shape a genuine misallocation usually
     leaves behind (see _split_across_codes).
  2. Whether this year's swing has happened before for THIS client - a
     recurring seasonal pattern reads very differently from a first-time
     anomaly, but a single job's current-vs-comparative check has no way
     to know which one it's looking at. Solved by keeping a small
     year-on-year history per client (client["pl_history"], written by
     snapshot_into_history after each run) that this year's run is
     checked against (see annotate_with_history) - accumulates over time
     as more jobs are generated for the same client, so the very first
     year has no history to compare against and every later one has
     progressively more.
  3. The actual "P&L Notes" a preparer would build by hand: every
     contact posting to EVERY P&L code (not just flagged ones) with its
     own current/previous £ AND the actual GL narrative behind each figure
     (deduplicated, current/previous year kept separate - see
     _extract_description), a per-code main variance driver (the single
     contact whose own movement explains most of that code's swing), and
     a duplicate-name flag - a contact who posted to more than one P&L
     code this year, checked across the whole P&L regardless of whether
     either code happened to breach the variance threshold (unlike
     _split_across_codes, which only looks at contacts already tied to a
     flagged account). See _pl_notes_detail.

All three are advisory, same as every other check in this system - they
surface candidates and context for a human reviewer to judge, never an
automatic accept/reject. See main.py's /jobs/{id}/pl-variance/run route
and the client-facing report route for how this gets used.
"""
from __future__ import annotations

import io

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from app.recon import MATERIALITY_AMOUNT, VARIANCE_PCT_THRESHOLD, ReconResult

NAME = "P&L variance review (current vs comparative)"


def _account_key(account_code, account_name) -> str:
    code = "" if account_code is None or (isinstance(account_code, float) and pd.isna(account_code)) else str(account_code).strip()
    return f"{code}|{account_name}"


def _split_across_codes(nominal_current: pd.DataFrame, flagged_keys: set[str]) -> pd.DataFrame:
    """For every contact who posted to a flagged P&L account this year,
    shows every nominal code (flagged or not) that same contact also
    posted to - a contact split across several codes is exactly the shape
    a genuine misallocation leaves behind, but it's just as often a
    contact whose spend genuinely spans several real categories (a
    supplier billing both repairs AND consumables, say). Left for the
    reviewer to judge which; a contact who only ever posted to the one
    flagged account isn't shown at all, since there's nothing to compare."""
    if nominal_current is None or nominal_current.empty or "contact" not in nominal_current.columns:
        return pd.DataFrame()
    df = nominal_current.copy()
    df["amount"] = df["debit"].fillna(0.0) - df["credit"].fillna(0.0)
    df["key"] = [_account_key(c, n) for c, n in zip(df["account_code"], df["account_name"])]
    df = df[df["contact"].notna() & (df["contact"].astype(str).str.strip() != "")]

    contacts_of_interest = df.loc[df["key"].isin(flagged_keys), "contact"].unique()
    if len(contacts_of_interest) == 0:
        return pd.DataFrame()

    rows = []
    for contact in contacts_of_interest:
        g = df[df["contact"] == contact]
        by_code = g.groupby(["account_code", "account_name"]).agg(postings=("amount", "count"), total=("amount", "sum")).reset_index()
        if len(by_code) < 2:
            continue  # this contact only ever posted to the flagged account - nothing to split
        for _, r in by_code.iterrows():
            key = _account_key(r["account_code"], r["account_name"])
            rows.append({
                "Contact": contact,
                "Account": f"{r['account_code']} - {r['account_name']}" if r["account_code"] else r["account_name"],
                "Postings": int(r["postings"]),
                "Total": round(float(r["total"]), 2),
                "This is the flagged account": key in flagged_keys,
            })
    if not rows:
        return pd.DataFrame()
    out = pd.DataFrame(rows)
    return out.sort_values(["Contact", "This is the flagged account"], ascending=[True, False]).reset_index(drop=True)


def _code_text(value) -> str:
    return "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value).strip()


def _natural_code_key(code: str) -> tuple:
    """Sorts account codes numerically when they parse as numbers (a real
    chart of accounts reads "90" before "100"), falling back to plain text
    for alpha-numeric codes - used only for _pl_notes_detail's own output
    order (a preparer expects a P&L note in ascending code order), never
    for pl_variance_analysis's `merged` table, which stays ordered by
    variance size (biggest movers first) for review priority."""
    try:
        return (0, float(code))
    except ValueError:
        return (1, code)


def _extract_description(raw: str, contact: str) -> str:
    """Turns a raw GL narrative into the short, non-redundant text worth
    showing next to a contact's own P&L note line - preferring whatever
    comes after the first "-" (most ledger exports prefix a boilerplate
    reference/contact before the actual narrative, e.g. "INV1023 - office
    repairs"), and failing that, stripping the contact's own name off the
    front (e.g. "Acme Ltd: office repairs") since restating the contact
    that's already the row's own label adds nothing."""
    description = raw.strip()
    if not description:
        return ""

    dash_index = description.find("-")
    if dash_index != -1:
        after_dash = description[dash_index + 1:].strip()
        if after_dash:
            return after_dash

    contact = contact.strip()
    if contact and description.lower().startswith(contact.lower()):
        description = description[len(contact):].lstrip(" :|/").strip()

    return description


def _ordered_unique_descriptions(df: pd.DataFrame) -> dict[tuple[str, str], str]:
    """Groups a prepared nominal-activity frame by (account_code, contact)
    into one semicolon-joined description per group - first-seen order,
    duplicates dropped - the shape both the current-year and previous-year
    description columns need. Returns {} for an empty frame or one with no
    description column at all (a generic-mapped upload that never mapped
    one), so a missing description never breaks the rest of the note."""
    if df.empty or "description" not in df.columns:
        return {}
    seen: dict[tuple[str, str], list[str]] = {}
    for code, contact, raw in zip(df["account_code"], df["contact"], df["description"]):
        text = _extract_description(_code_text(raw), contact)
        if not text:
            continue
        bucket = seen.setdefault((code, contact), [])
        if text not in bucket:
            bucket.append(text)
    return {key: "; ".join(values) for key, values in seen.items()}


def _pl_notes_detail(pl_summary: pd.DataFrame, nominal_current: pd.DataFrame, nominal_comparative: pd.DataFrame | None) -> pd.DataFrame:
    """Builds the supplier/customer-level "P&L Notes" drill-down described
    in this module's own docstring (point 3): every contact posting to
    every P&L code this year (and last year too, when the practice has
    also uploaded prior-year nominal activity - a comparative-year contact
    breakdown isn't available from pl_comparative alone, which only ever
    carries account-level totals), each with its own current/previous-year
    GL narrative alongside the £ (see _extract_description -
    deduplicated, blank on the TOTAL row since it has no single narrative
    of its own), a per-code TOTAL row carrying that code's own main
    variance driver, and a universal duplicate-name flag.

    `pl_summary` is pl_variance_analysis's own already-computed `merged`
    frame (account_code/account_name/current_year/comparative_year/
    variance_pct/flag), reused rather than recomputed, so this table's
    own numbers can never drift from the headline review table's for the
    same code. Without a prior-year nominal upload, every "Previous Year"
    contact figure here is 0 even though the code's own TOTAL row still
    shows the real comparative_year total from pl_summary - the gap is
    visible to the reviewer rather than silently papered over."""
    if pl_summary is None or pl_summary.empty:
        return pd.DataFrame()
    if nominal_current is None or nominal_current.empty or "contact" not in nominal_current.columns:
        return pd.DataFrame()

    def _prepare(df: pd.DataFrame) -> pd.DataFrame:
        out = df.copy()
        out["amount"] = out["debit"].fillna(0.0) - out["credit"].fillna(0.0)
        out["contact"] = out["contact"].fillna("").astype(str).str.strip()
        out["account_code"] = out["account_code"].map(_code_text)
        return out[out["contact"] != ""]

    cur = _prepare(nominal_current)
    comp = _prepare(nominal_comparative) if nominal_comparative is not None and not nominal_comparative.empty and "contact" in nominal_comparative.columns else pd.DataFrame(columns=cur.columns)

    current_descriptions = _ordered_unique_descriptions(cur)
    previous_descriptions = _ordered_unique_descriptions(comp)

    pl_codes = {c for c in pl_summary["account_code"].map(_code_text) if c}

    # Universal duplicate-name check: every OTHER P&L code this contact has
    # also posted to THIS YEAR, across the whole P&L range - not limited to
    # contacts tied to an already-flagged account the way
    # _split_across_codes' equivalent check is.
    contact_codes: dict[str, set[str]] = {}
    for contact, code in zip(cur["contact"], cur["account_code"]):
        if code in pl_codes:
            contact_codes.setdefault(contact, set()).add(code)

    rows = []
    ordered_summary = pl_summary.assign(_code_text=pl_summary["account_code"].map(_code_text))
    ordered_summary = ordered_summary.sort_values("_code_text", key=lambda s: s.map(_natural_code_key))

    for _, head in ordered_summary.iterrows():
        code = head["_code_text"]
        account_label = f"{code} - {head['account_name']}" if code else head["account_name"]

        current_by_contact = cur.loc[cur["account_code"] == code].groupby("contact")["amount"].sum()
        previous_by_contact = comp.loc[comp["account_code"] == code].groupby("contact")["amount"].sum() if not comp.empty else pd.Series(dtype=float)

        contacts = sorted(set(current_by_contact.index) | set(previous_by_contact.index))
        if not contacts:
            continue

        main_driver, main_movement = None, 0.0
        for contact in contacts:
            movement = float(current_by_contact.get(contact, 0.0)) - float(previous_by_contact.get(contact, 0.0))
            if main_driver is None or abs(movement) > abs(main_movement):
                main_driver, main_movement = contact, movement

            duplicate = len(contact_codes.get(contact, set())) > 1
            rows.append({
                "Account": account_label,
                "Contact": contact,
                "Current Period Description": current_descriptions.get((code, contact), ""),
                "Previous Period Description": previous_descriptions.get((code, contact), ""),
                "Current Year": round(float(current_by_contact.get(contact, 0.0)), 2),
                "Previous Year": round(float(previous_by_contact.get(contact, 0.0)), 2),
                "Variance %": "",  # only meaningful at code level - see the TOTAL row below
                "Duplicate Name": "Booked to more than one P&L head in current year" if duplicate else "",
                "Main Variance Driver": "",
            })

        driver_text = ""
        if bool(head.get("flag")) and main_driver is not None:
            direction = "Higher" if main_movement > 0 else "Lower"
            sign = "+" if main_movement >= 0 else "-"
            driver_text = f"{direction} - {main_driver} ({sign}£{abs(main_movement):,.2f})"

        rows.append({
            "Account": account_label,
            "Contact": "TOTAL",
            "Current Period Description": "",
            "Previous Period Description": "",
            "Current Year": round(float(head["current_year"]), 2),
            "Previous Year": round(float(head["comparative_year"]), 2),
            "Variance %": round(float(head["variance_pct"]) * 100, 1),
            "Duplicate Name": "",
            "Main Variance Driver": driver_text,
        })

    return pd.DataFrame(rows)


def pl_variance_analysis(
    pl_current: pd.DataFrame, pl_comparative: pd.DataFrame, nominal_current: pd.DataFrame,
    nominal_comparative: pd.DataFrame | None = None,
    materiality: float = MATERIALITY_AMOUNT, variance_pct_threshold: float = VARIANCE_PCT_THRESHOLD,
) -> ReconResult:
    if pl_current is None or pl_current.empty:
        return ReconResult(NAME, "n/a", "No current-year Profit & Loss uploaded (or derivable from a Trial Balance).")

    cur = pl_current.groupby(["account_code", "account_name", "category"], as_index=False, dropna=False)["amount"].sum()
    cur = cur.rename(columns={"amount": "current_year"})
    if pl_comparative is not None and not pl_comparative.empty:
        comp = pl_comparative.groupby(["account_code", "account_name"], as_index=False)["amount"].sum()
        comp = comp.rename(columns={"amount": "comparative_year"})
    else:
        comp = pd.DataFrame(columns=["account_code", "account_name", "comparative_year"])

    merged = pd.merge(cur, comp, on=["account_code", "account_name"], how="outer")
    # An account with no category mapped comes through from apply_mapping
    # as an empty string, not NaN (it fills every unmapped non-numeric
    # canonical field with "" - see parsers.apply_mapping), so a bare
    # fillna() alone doesn't catch it; a comparative-only row (no matching
    # current-year account) genuinely has no category at all, which IS
    # NaN, since it never went through cur's category column - both need
    # the same fallback label.
    merged["category"] = merged["category"].replace("", pd.NA).fillna("Unclassified")
    merged[["current_year", "comparative_year"]] = merged[["current_year", "comparative_year"]].fillna(0.0)
    merged["variance_amount"] = merged["current_year"] - merged["comparative_year"]
    merged["variance_pct"] = merged.apply(
        lambda r: (r["variance_amount"] / abs(r["comparative_year"])) if r["comparative_year"] not in (0, 0.0) else (1.0 if r["variance_amount"] != 0 else 0.0),
        axis=1,
    )
    merged["flag"] = (merged["variance_amount"].abs() >= materiality) & (merged["variance_pct"].abs() >= variance_pct_threshold)
    merged = merged.sort_values("variance_amount", key=lambda s: s.abs(), ascending=False).reset_index(drop=True)

    flagged = merged[merged["flag"]]
    status = "ok" if flagged.empty else "review"
    msg = (
        "No P&L account moved beyond materiality year-on-year."
        if flagged.empty else
        f"{len(flagged)} P&L account(s) moved >{variance_pct_threshold:.0%} and >£{materiality:,.0f} year-on-year - "
        f"review for reclassification, genuine business change, or a coding slip."
    )

    result = ReconResult(NAME, status, msg, merged)
    if not flagged.empty:
        flagged_keys = {_account_key(r["account_code"], r["account_name"]) for _, r in flagged.iterrows()}
        split_detail = _split_across_codes(nominal_current, flagged_keys)
        if not split_detail.empty:
            result.extra_detail = split_detail
            result.extra_detail_label = "Contacts posting to a flagged account AND at least one other code this year"

    notes_detail = _pl_notes_detail(merged, nominal_current, nominal_comparative)
    if not notes_detail.empty:
        result.matched_detail = notes_detail
        result.matched_detail_label = "P&L Notes - supplier/customer analysis by nominal code"
    return result


def annotate_with_history(detail: pd.DataFrame, history: dict) -> pd.DataFrame:
    """Adds a "Historical pattern" column to a pl_variance_analysis
    result's own detail table: for each account, looks at the client's
    stored prior-year snapshots (see snapshot_into_history) and says
    whether this year's swing (or lack of one) has been seen before -
    purely descriptive, for the reviewer to weigh, not an automatic
    accept/reject. A brand-new client with no prior jobs generated yet
    simply gets "No prior-year history yet" on every row."""
    if detail is None or detail.empty or not history:
        if detail is not None and not detail.empty:
            detail = detail.copy()
            detail["historical_pattern"] = "No prior-year history yet"
        return detail

    def _pattern(row) -> str:
        key = _account_key(row["account_code"], row["account_name"])
        prior_years = [snap[key] for snap in history.values() if key in snap]
        if not prior_years:
            return "No prior-year history yet"
        prior_flagged = sum(1 for y in prior_years if y.get("flagged"))
        if not row["flag"]:
            return f"Not flagged this year ({len(prior_years)} prior year(s) on record)"
        if prior_flagged:
            return f"Recurring - flagged in {prior_flagged} of {len(prior_years)} prior year(s) too - may be normal for this client"
        return f"New this year - not flagged in {len(prior_years)} prior year(s) on record"

    detail = detail.copy()
    detail["historical_pattern"] = detail.apply(_pattern, axis=1)
    return detail


def snapshot_into_history(client: dict, year_end: str, variance_detail: pd.DataFrame) -> None:
    """Records this year's own P&L account figures into the client's
    year-on-year history (client["pl_history"], keyed by accounting
    year-end so re-running the same job's check overwrites that year's
    own snapshot rather than accumulating duplicates) - mutates client in
    place; the caller still needs storage.save_client(client) to persist
    it. Every account is stored, not just flagged ones, so a quiet
    account this year is still on record as a comparison point for next
    year's run."""
    if variance_detail is None or variance_detail.empty or not year_end:
        return
    history = client.setdefault("pl_history", {})
    snapshot = {}
    for _, row in variance_detail.iterrows():
        key = _account_key(row["account_code"], row["account_name"])
        snapshot[key] = {
            "account_name": row["account_name"],
            "amount": float(row["current_year"]),
            "variance_amount": float(row["variance_amount"]),
            "variance_pct": float(row["variance_pct"]),
            "flagged": bool(row["flag"]),
        }
    history[year_end] = snapshot


def _plain_comment(row: pd.Series) -> str:
    """Composes the client-facing comment for one P&L line - plain
    English only, no internal workpaper language ("flag", "materiality",
    "review status"). A line that didn't move meaningfully gets no
    comment at all rather than a hollow "no issues" filler on every row."""
    if not row["flag"]:
        return ""
    direction = "increased" if row["variance_amount"] > 0 else "decreased"
    pct = abs(row["variance_pct"]) * 100
    comment = f"{direction.capitalize()} by {abs(row['variance_amount']):,.0f} ({pct:,.0f}%) compared to last year."
    pattern = str(row.get("historical_pattern", ""))
    if pattern.startswith("Recurring"):
        comment += " Similar movement seen in prior years for this client."
    elif pattern.startswith("New this year"):
        comment += " First time this account has moved by this much."
    return comment


_CLIENT_REPORT_HEADER_FILL = PatternFill("solid", fgColor="1F3864")
_CLIENT_REPORT_HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
_CLIENT_REPORT_TITLE_FONT = Font(name="Calibri", size=14, bold=True)
_CLIENT_REPORT_SUBTITLE_FONT = Font(name="Calibri", size=11, italic=True, color="595959")


def build_client_report(client_name: str, period_label: str, variance_detail: pd.DataFrame) -> bytes:
    """A separate, deliberately simpler export of the same variance review
    meant to go straight to the client - every P&L account for the
    period, in normal P&L order (category, then account name), with a
    plain-English comment only on the lines that moved meaningfully. No
    internal review status, materiality threshold, or nominal-code
    investigation detail - see main.py's client-report download route
    for the reasoning."""
    wb = Workbook()
    ws = wb.active
    ws.title = "P&L Variance Summary"

    ws["A1"] = client_name
    ws["A1"].font = _CLIENT_REPORT_TITLE_FONT
    ws["A2"] = f"Profit & Loss - Year-on-Year Variance Summary - {period_label}"
    ws["A2"].font = _CLIENT_REPORT_SUBTITLE_FONT

    headers = ["Category", "Account", "This Year", "Last Year", "Variance", "Variance %", "Comment"]
    header_row = 4
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=header)
        cell.font = _CLIENT_REPORT_HEADER_FONT
        cell.fill = _CLIENT_REPORT_HEADER_FILL

    if variance_detail is None or variance_detail.empty:
        ws.cell(row=header_row + 2, column=1, value="No Profit & Loss data available for this period.")
    else:
        ordered = variance_detail.sort_values(["category", "account_name"]).reset_index(drop=True)
        row = header_row + 1
        for _, r in ordered.iterrows():
            ws.cell(row=row, column=1, value=r["category"])
            ws.cell(row=row, column=2, value=r["account_name"])
            ws.cell(row=row, column=3, value=round(float(r["current_year"]), 2)).number_format = "#,##0"
            ws.cell(row=row, column=4, value=round(float(r["comparative_year"]), 2)).number_format = "#,##0"
            ws.cell(row=row, column=5, value=round(float(r["variance_amount"]), 2)).number_format = "#,##0;(#,##0)"
            ws.cell(row=row, column=6, value=round(float(r["variance_pct"]), 4)).number_format = "0%"
            comment_cell = ws.cell(row=row, column=7, value=_plain_comment(r))
            comment_cell.alignment = Alignment(wrap_text=True)
            row += 1

    widths = [16, 28, 13, 13, 13, 11, 60]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + col)].width = width

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
