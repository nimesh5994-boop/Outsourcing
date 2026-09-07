"""Unit tests for the VAT filing-period wizard's pure logic - see
app/vat_periods.py. No database needed."""
from datetime import date

import pytest

from app import vat_periods


def test_expected_period_ends_quarterly_covers_a_full_accounting_year():
    ends = vat_periods.expected_period_ends("2025-03-01", "2026-02-28", "quarterly")
    assert ends == [
        date(2025, 5, 31), date(2025, 8, 31), date(2025, 11, 30), date(2026, 2, 28),
    ]


def test_expected_period_ends_monthly_covers_a_full_accounting_year():
    ends = vat_periods.expected_period_ends("2025-01-01", "2025-12-31", "monthly")
    assert len(ends) == 12
    assert ends[0] == date(2025, 1, 31)
    assert ends[-1] == date(2025, 12, 31)


def test_expected_period_ends_handles_a_short_first_period():
    """A new client's first job can be a part-year period - the walk-back
    should still land sensibly without assuming a clean 4-quarter split.
    Walking backward from the year-end in exact 3-month steps (rather than
    forward from the job's own start) also matches how a real VAT quarter
    scheme actually works: which calendar months a client's quarters cover
    is fixed by their HMRC registration, not by when this particular job's
    accounting period happens to start - so the first quarter here is a
    genuine one-month stub (01-30 Jun), not a padded-out three-month one."""
    ends = vat_periods.expected_period_ends("2025-06-01", "2025-12-31", "quarterly")
    assert ends == [date(2025, 6, 30), date(2025, 9, 30), date(2025, 12, 31)]


def test_expected_period_ends_rejects_unknown_scheme():
    with pytest.raises(ValueError, match="Unknown VAT period type"):
        vat_periods.expected_period_ends("2025-01-01", "2025-12-31", "fortnightly")


def test_period_coverage_matches_within_tolerance_and_flags_missing():
    expected = [date(2025, 5, 31), date(2025, 8, 31), date(2025, 11, 30), date(2026, 2, 28)]
    detected = [date(2025, 5, 29), date(2025, 8, 31), None]  # one slightly off, one exact, one unparsed
    coverage = vat_periods.period_coverage(expected, detected)
    assert [c["covered"] for c in coverage] == [True, True, False, False]


def test_period_coverage_does_not_let_one_upload_satisfy_two_expected_periods():
    """Two quarters both uploaded as the same (wrong) detected date
    shouldn't silently mark both as covered - only one expected period
    can be satisfied per actual uploaded file."""
    expected = [date(2025, 5, 31), date(2025, 8, 31)]
    detected = [date(2025, 5, 31)]
    coverage = vat_periods.period_coverage(expected, detected)
    assert [c["covered"] for c in coverage] == [True, False]
