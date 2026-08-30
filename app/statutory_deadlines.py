"""Statutory filing deadlines - purely deterministic, computed from the
job's own period dates alone (no upload needed), yet completely absent
from the working paper today. A real firm's job file always states these
somewhere; this system had nowhere for them at all.

Three deadlines for an ordinary UK private limited company:

  - Companies House annual accounts filing deadline: normally 9 months
    after the Accounting Reference Date (period end). Governed by
    Companies Act 2006 s.442's ARD special case - if the period end is
    itself the last day of its calendar month, the deadline is the last
    day of the target month N months later (regardless of that month's
    own length), not just the same day-of-month; see _add_months_ch_style.
  - Corporation Tax return (CT600) filing deadline: 12 months after the
    end of the accounting period - HMRC's rule is the plain anniversary
    date, no ARD special case.
  - Corporation Tax payment deadline: 9 months and 1 day after the end
    of the accounting period, for a company NOT paying by quarterly
    instalments (broadly, one with augmented profits under £1.5m,
    scaled for a short period or associated companies) - stated as an
    explicit assumption in the message, since a large company's real
    payment dates are quarterly instalments this system has no way to
    compute without knowing its instalment-payment status.

Deliberately doesn't attempt: first-year-of-trading accounts (a company's
first accounts can cover more than 12 months, with its own 21-months-
from-incorporation rule requiring an incorporation date this system
doesn't capture), or the confirmation statement (due from the
incorporation/last-statement anniversary, not the accounting period, so
not derivable from period dates at all).
"""
from dataclasses import dataclass
from datetime import date, timedelta

import pandas as pd
from dateutil.relativedelta import relativedelta

from app.recon import ReconResult

LARGE_COMPANY_INSTALMENT_THRESHOLD = 1_500_000.0  # £ augmented profits - above this, quarterly instalments apply instead


@dataclass
class StatutoryDeadlines:
    period_end: date
    accounts_filing_deadline: date
    ct_return_filing_deadline: date
    ct_payment_deadline: date


def _add_months_ch_style(d: date, months: int) -> date:
    """Companies House's ARD rule: if `d` is the last day of its calendar
    month, the result is the last day of the target month `months` later
    - regardless of how many days that month actually has - not just the
    same day-of-month. dateutil's relativedelta(day=31) clamps to a
    month's real last day, which is exactly this rule; the plain
    same-day-of-month case is relativedelta's own default behaviour."""
    is_last_day_of_month = (d + timedelta(days=1)).month != d.month
    if is_last_day_of_month:
        return d + relativedelta(months=months, day=31)
    return d + relativedelta(months=months)


def compute(period_end: date) -> StatutoryDeadlines:
    return StatutoryDeadlines(
        period_end=period_end,
        accounts_filing_deadline=_add_months_ch_style(period_end, 9),
        ct_return_filing_deadline=period_end + relativedelta(months=12),
        ct_payment_deadline=period_end + relativedelta(months=9) + timedelta(days=1),
    )


def summary_message(deadlines: StatutoryDeadlines) -> str:
    return (
        f"Companies House accounts filing deadline: {deadlines.accounts_filing_deadline:%d %B %Y}. "
        f"Corporation Tax return (CT600) filing deadline: {deadlines.ct_return_filing_deadline:%d %B %Y}. "
        f"Corporation Tax payment deadline: {deadlines.ct_payment_deadline:%d %B %Y} - assumes the company "
        f"pays by a single instalment (augmented profits under £1.5m; scale down for a short period or "
        f"associated companies), not the quarterly instalment regime a large company must use instead. "
        f"Doesn't cover a first set of accounts (its own 21-months-from-incorporation rule) or the "
        f"confirmation statement (due from the incorporation/last-statement anniversary, not the "
        f"accounting period) - neither is derivable from the job's period dates alone."
    )


def build_result(period_end: date | None) -> ReconResult:
    """Always "ok" - these are facts, not findings, so there's nothing to
    flag. Still worth a result in the pipeline: it's the only way these
    dates land on the Index sheet and every schedule's summary table
    without a preparer having to work them out by hand."""
    name = "Statutory filing deadlines"
    if period_end is None:
        return ReconResult(name, "n/a", "No period end date available to compute filing deadlines from.")

    deadlines = compute(period_end)
    detail = pd.DataFrame([
        {"Deadline": "Companies House - annual accounts filing", "Date": deadlines.accounts_filing_deadline},
        {"Deadline": "HMRC - Corporation Tax return (CT600) filing", "Date": deadlines.ct_return_filing_deadline},
        {"Deadline": "HMRC - Corporation Tax payment", "Date": deadlines.ct_payment_deadline},
    ])
    return ReconResult(name, "ok", summary_message(deadlines), detail)
