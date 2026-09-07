"""Auto-generates the VAT filing periods a job should expect from a single
one-time choice of VAT scheme (quarterly/monthly), so a preparer doesn't
have to work out which quarters/months to expect and chase each one down
by hand. See main.py's vat-setup route and the VAT Return upload section
in job_detail.html, which shows each expected period as covered/missing
against what's actually been uploaded (matched by each upload's own
detected period-end date - see document_detection.guess_period and
xero_reports.extract_period_info)."""
from __future__ import annotations

from datetime import date

import pandas as pd

VAT_PERIOD_TYPES = {"quarterly": 3, "monthly": 1}
VAT_PERIOD_TYPE_LABELS = {"quarterly": "Quarterly", "monthly": "Monthly"}


def expected_period_ends(period_start: str | date, period_end: str | date, vat_period_type: str) -> list[date]:
    """Every period-end date a client on this VAT scheme should have filed
    by, walking backward from the accounting year-end in 3-month
    (quarterly) or 1-month (monthly) steps until reaching the year-start -
    so a part-year first job (e.g. a new client's first 10-month period)
    still gets a sensible, non-overlapping set of expected filing dates
    rather than assuming a clean 4/12 split. MonthEnd arithmetic (not a
    flat 90/30-day subtraction) so a leap-February year-end still steps
    back to genuine month-end dates (31 May, not 28 May)."""
    months = VAT_PERIOD_TYPES.get(vat_period_type)
    if months is None:
        raise ValueError(f"Unknown VAT period type: {vat_period_type!r}")
    start = pd.Timestamp(period_start)
    end = pd.Timestamp(period_end)

    ends: list[date] = []
    cursor = end
    while cursor >= start:
        ends.append(cursor.date())
        cursor = cursor - pd.offsets.MonthEnd(months)
    ends.reverse()
    return ends


def period_coverage(
    expected_ends: list[date], detected_ends: list[date | None], tolerance_days: int = 5,
) -> list[dict]:
    """Matches each expected period end against the nearest confirmed
    upload's own detected period-end date (within tolerance_days, since a
    real export's own date is exact but preparers can still be a day or
    two off manually keying period boundaries in the job setup) and
    reports which expected periods are covered and which are still
    missing. Each detected end can only satisfy one expected period, so
    two uploads of the same quarter don't silently mask a genuinely
    missing one."""
    remaining = list(detected_ends)
    coverage = []
    for expected in expected_ends:
        match = next((d for d in remaining if d is not None and abs((d - expected).days) <= tolerance_days), None)
        if match is not None:
            remaining.remove(match)
        coverage.append({"expected_end": expected, "covered": match is not None})
    return coverage
