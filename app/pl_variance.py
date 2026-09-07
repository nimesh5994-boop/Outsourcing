"""P&L variance review: current vs comparative movement on every P&L
account, with two pieces of extra context a bare TB-wide variance check
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

Both pieces are advisory, same as every other check in this system - they
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


def pl_variance_analysis(
    pl_current: pd.DataFrame, pl_comparative: pd.DataFrame, nominal_current: pd.DataFrame,
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
