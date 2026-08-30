"""Unit tests for the statutory filing deadline calculator - see
app/statutory_deadlines.py. Purely deterministic date arithmetic (no
data upload involved), so these pin down the exact UK rules rather than
general behaviour:

  - Companies House's Accounting Reference Date (ARD) special case: if
    the period end is the last day of its calendar month, the accounts
    filing deadline is the last day of the target month N months later
    (regardless of that month's own length) - not just the same
    day-of-month.
  - HMRC's CT600 filing deadline and CT payment deadline don't have that
    special case - they're a plain anniversary date / a fixed
    9-months-and-a-day offset.
"""
from datetime import date

from app.statutory_deadlines import build_result, compute


def test_accounts_deadline_last_day_of_december_goes_to_last_day_of_september():
    # the single most common UK year end - widely known real-world fact
    d = compute(date(2025, 12, 31))
    assert d.accounts_filing_deadline == date(2026, 9, 30)


def test_accounts_deadline_last_day_of_february_non_leap_year():
    # 28 Feb is the last day of Feb in a non-leap year - the ARD rule
    # means the deadline is the LAST day of November (30th), not the 28th
    d = compute(date(2025, 2, 28))
    assert d.accounts_filing_deadline == date(2025, 11, 30)


def test_accounts_deadline_not_last_day_of_month_uses_same_day_number():
    d = compute(date(2025, 6, 15))
    assert d.accounts_filing_deadline == date(2026, 3, 15)


def test_accounts_deadline_last_day_of_january_lands_on_31_october():
    # Jan 31 is the last day of January; October also has 31 days, so the
    # "last day of the target month" rule and "same day number" rule
    # happen to coincide here - still worth pinning down explicitly
    d = compute(date(2025, 1, 31))
    assert d.accounts_filing_deadline == date(2025, 10, 31)


def test_ct600_filing_deadline_is_plain_twelve_month_anniversary():
    d = compute(date(2025, 12, 31))
    assert d.ct_return_filing_deadline == date(2026, 12, 31)


def test_ct_payment_deadline_is_nine_months_and_one_day():
    d = compute(date(2025, 12, 31))
    assert d.ct_payment_deadline == date(2026, 10, 1)


def test_build_result_is_always_ok_with_all_three_dates_in_detail():
    result = build_result(date(2025, 12, 31))
    assert result.status == "ok"
    assert set(result.detail["Date"]) == {date(2026, 9, 30), date(2026, 12, 31), date(2026, 10, 1)}
    assert "30 September 2026" in result.message


def test_build_result_na_without_a_period_end():
    assert build_result(None).status == "n/a"
