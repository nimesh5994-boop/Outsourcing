"""Best-effort auto-detection of what an uploaded file actually is, so
staff don't have to pick report type / platform / period from dropdowns
before uploading - they just drop files in, and the system guesses, then
shows the guess for a quick confirm (see the "Auto-detect + confirm"
rationale in the upload flow) rather than silently trusting a heuristic
with real client financials on the line.

Three independent guesses, in order of reliability:
  1. report_type + platform: try each Xero-native parser first (if one
     succeeds without error, that's a near-certain signal - both the
     report type AND that it's a genuine Xero export); otherwise score
     column headers against every report type's alias dictionary and
     take the best match, platform falls back to a light column-name
     heuristic (mostly "other", since only Xero has a dedicated parser).
  2. period (current vs comparative): reuse the existing Xero title-row
     period extraction where available; otherwise look for the latest
     date in any date-shaped column and compare it to the job's declared
     period-end dates; if the file has no dates at all (a point-in-time
     report like a TB or aged listing), fall back to "if this report
     type already has a confirmed 'current' upload on this job, this new
     one is probably the comparative" - matches the common real workflow
     of uploading this year's export, then last year's.
None of these guesses are ever auto-accepted without a human seeing them
first - see the confirm step in main.py's upload route.
"""
from __future__ import annotations

import pandas as pd

from app import mapping, xero_reports
from app.models import REPORT_SCHEMAS, REPORT_TYPES, REQUIRED_FIELDS
from app.parsers import DataSource

XERO_NATIVE_REPORT_TYPES = {"trial_balance", "nominal_activity", "aged_debtors", "aged_creditors"}

_PLATFORM_HINTS = {
    "sage": {"nominalcode", "nominal", "sagereference"},
    "qbo": {"class", "quickbooksid", "qbid"},
}


def try_xero_native(source: DataSource) -> str | None:
    """Attempts each Xero-native parser in turn; returns the report_type of
    the first one that parses cleanly, or None if none of them match this
    file's shape at all. A successful native parse is a much stronger
    signal than column-name guessing, since it means the file's actual
    Xero-specific layout (title rows, grouped sections, embedded
    comparative column) was recognised, not just similar-looking headers."""
    for report_type in XERO_NATIVE_REPORT_TYPES:
        try:
            if report_type == "trial_balance":
                xero_reports.parse_trial_balance(source)
            elif report_type == "nominal_activity":
                xero_reports.parse_account_transactions(source)
            elif report_type == "aged_debtors":
                xero_reports.parse_aged_report(source, "customer")
            elif report_type == "aged_creditors":
                xero_reports.parse_aged_report(source, "supplier")
            return report_type
        except Exception:
            continue
    return None


def classify_report_type(columns: list[str]) -> tuple[str | None, float]:
    """Scores every report type's alias dictionary against these column
    headers; returns (best_report_type, confidence 0-1) or (None, 0.0) if
    nothing matched well enough to guess. Required fields count double,
    since a report missing its own required fields isn't a real match
    even if a couple of incidental columns happen to line up."""
    scores: dict[str, float] = {}
    for report_type in REPORT_TYPES:
        suggestion = mapping.suggest_mapping(report_type, columns)
        matched_fields = {v for v in suggestion.values() if v}
        if not matched_fields:
            continue
        required = set(REQUIRED_FIELDS.get(report_type, []))
        score = len(matched_fields) + len(matched_fields & required) * 2
        scores[report_type] = score

    if not scores:
        return None, 0.0

    best_type = max(scores, key=scores.get)
    best_score = scores[best_type]

    # profit_and_loss and balance_sheet share an identical column shape
    # (account_code/account_name/category/amount) - the only way to tell
    # them apart is what the category *values* actually say.
    pl_score, bs_score = scores.get("profit_and_loss", 0), scores.get("balance_sheet", 0)
    if pl_score and bs_score and pl_score == bs_score == best_score:
        best_type = "profit_and_loss" if pl_score >= bs_score else "balance_sheet"

    possible_max = len(REPORT_SCHEMAS.get(best_type, {})) + len(REQUIRED_FIELDS.get(best_type, [])) * 2
    confidence = min(1.0, best_score / max(possible_max, 1))
    return best_type, confidence


def disambiguate_pl_vs_bs(df: pd.DataFrame, category_column: str | None) -> str:
    """Once column shape alone can't tell P&L from B/S apart, peek at what
    the category values actually say."""
    if category_column is None or category_column not in df.columns:
        return "profit_and_loss"
    values = " ".join(str(v).lower() for v in df[category_column].dropna().unique())
    bs_hits = sum(kw in values for kw in ("asset", "liabilit", "equity", "capital", "reserve"))
    pl_hits = sum(kw in values for kw in ("turnover", "sales", "cost of sales", "overhead", "expense", "income"))
    return "balance_sheet" if bs_hits > pl_hits else "profit_and_loss"


def classify_platform(columns: list[str], is_xero_native: bool) -> str:
    if is_xero_native:
        return "xero"
    normalised = {mapping._normalise(c) for c in columns}
    for platform, hints in _PLATFORM_HINTS.items():
        if normalised & hints:
            return platform
    return "other"


_DATE_LIKE_FIELDS = {"date", "statement_date", "date_acquired"}


def _latest_date_in_columns(source: DataSource, report_type: str) -> "pd.Timestamp | None":
    """For report types with a per-row date (nominal activity, bank
    statement, fixed asset register), find whichever source column maps to
    a date field and return the latest value in it - a real signal of
    which period the file covers. Reports with no date column at all (TB,
    aged listings, VAT return) return None; guess_period() falls back to
    the "second upload of this type = comparative" heuristic for those."""
    columns = source.raw_columns()
    suggestion = mapping.suggest_mapping(report_type, columns)
    date_columns = [col for col, field in suggestion.items() if field in _DATE_LIKE_FIELDS]
    if not date_columns:
        return None
    raw = source.raw_dataframe()
    best = None
    for col in date_columns:
        # format="mixed": see the identical fix/comment in
        # parsers.apply_mapping - without it, a mixed-format or even a
        # single-column-but-multi-row ISO date series can have day/month
        # silently swapped on some rows.
        parsed = pd.to_datetime(raw[col], errors="coerce", dayfirst=True, format="mixed")
        col_max = parsed.max()
        if pd.notna(col_max) and (best is None or col_max > best):
            best = col_max
    return best


def guess_period(source: DataSource, report_type: str, job: dict, xero_native_report_type: str | None) -> str:
    from datetime import date as _date

    current_end = job.get("current_period_end")
    comparative_end = job.get("comparative_period_end")
    current_end = _date.fromisoformat(current_end) if current_end else None
    comparative_end = _date.fromisoformat(comparative_end) if comparative_end else None

    found_end = None
    if xero_native_report_type:
        info = xero_reports.extract_period_info(source)
        found_end = info.get("end")
    if found_end is None:
        latest = _latest_date_in_columns(source, report_type)
        if latest is not None:
            found_end = latest.date()

    if found_end is not None:
        current_delta = abs((found_end - current_end).days) if current_end else None
        comparative_delta = abs((found_end - comparative_end).days) if comparative_end else None
        if comparative_delta is not None and (current_delta is None or comparative_delta < current_delta):
            return "comparative"
        if current_delta is not None:
            return "current"

    # No date signal at all (TB, aged listings, VAT return): if a
    # confirmed "current" upload of this exact report type already exists
    # on this job, this new one is very likely the comparative-year file.
    existing_current = any(
        u["report_type"] == report_type and u["period"] == "current" and u["confirmed"]
        for u in job.get("uploads", {}).values()
    )
    return "comparative" if existing_current else "current"
