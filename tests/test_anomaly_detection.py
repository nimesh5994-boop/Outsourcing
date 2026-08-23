"""Tests for the cross-transaction anomaly checks: a contact whose spend is
lopsidedly coded to the wrong nominal account, duplicate postings, and
manual journals posted on a weekend."""
import pandas as pd
import pytest

from app import anomaly_detection as ad


def _row(date, code, name, contact, debit=0.0, credit=0.0, reference="", description="", source_type="Bill"):
    return {
        "date": pd.Timestamp(date), "account_code": code, "account_name": name, "reference": reference,
        "description": description, "contact": contact, "source_type": source_type, "debit": debit, "credit": credit,
    }


def test_contact_coding_consistency_flags_the_bt_style_pattern():
    rows = []
    # BT: 10 postings to Telephone, 2 to Light & Heat - the minority 2 should be flagged
    for i in range(10):
        rows.append(_row(f"2025-0{(i % 9) + 1}-05", "7502", "Telephone", "BT", debit=45.0, reference=f"BT-{i}"))
    rows.append(_row("2025-03-10", "7500", "Light & Heat", "BT", debit=45.0, reference="BT-MISC-1"))
    rows.append(_row("2025-07-10", "7500", "Light & Heat", "BT", debit=45.0, reference="BT-MISC-2"))

    # a contact genuinely split ~evenly across two codes - should NOT be flagged
    for i in range(4):
        rows.append(_row(f"2025-0{i + 1}-15", "5000", "Materials", "MULTI SUPPLIES LTD", debit=200.0, reference=f"MS-{i}"))
    for i in range(4):
        rows.append(_row(f"2025-0{i + 1}-20", "5100", "Plant Hire", "MULTI SUPPLIES LTD", debit=150.0, reference=f"MS-H-{i}"))

    # a contact with too few transactions to trust any pattern
    rows.append(_row("2025-01-01", "7502", "Telephone", "SMALL VENDOR", debit=10.0))
    rows.append(_row("2025-02-01", "7500", "Light & Heat", "SMALL VENDOR", debit=10.0))

    nominal = pd.DataFrame(rows)
    result = ad.contact_coding_consistency(nominal)

    assert result.status == "review"
    assert len(result.detail) == 2
    assert set(result.detail["Contact"]) == {"BT"}
    assert all(result.detail["Posted to"] == "7500 - Light & Heat")
    assert all(result.detail["Likely correct code"] == "7502 - Telephone")
    assert "MULTI SUPPLIES LTD" not in result.detail["Contact"].values
    assert "SMALL VENDOR" not in result.detail["Contact"].values


def test_contact_coding_consistency_ok_when_nothing_lopsided():
    rows = [_row(f"2025-0{i + 1}-01", "7502", "Telephone", "BT", debit=45.0) for i in range(6)]
    nominal = pd.DataFrame(rows)
    result = ad.contact_coding_consistency(nominal)
    assert result.status == "ok"
    assert result.detail.empty


def test_duplicate_transactions_flags_same_contact_date_amount():
    rows = [
        _row("2025-03-01", "5000", "Materials", "ACME LTD", debit=500.0, reference="INV-001"),
        _row("2025-03-01", "5000", "Materials", "ACME LTD", debit=500.0, reference="INV-002"),  # same contact/date/amount, different ref
        _row("2025-04-01", "5000", "Materials", "ACME LTD", debit=300.0, reference="INV-003"),  # unrelated, not a dupe
    ]
    nominal = pd.DataFrame(rows)
    result = ad.duplicate_transactions(nominal)

    assert result.status == "review"
    assert len(result.detail) == 2
    assert all("same contact, date and amount" in r.lower() for r in result.detail["Flag reason"])


def test_duplicate_transactions_flags_repeated_reference_on_same_code_and_amount():
    # same reference, code, and amount but a different date - a genuine
    # "entered twice by mistake" signal, distinct from the amount-based
    # check (which requires the same date)
    rows = [
        _row("2025-03-01", "5000", "Materials", "ACME LTD", debit=500.0, reference="INV-777"),
        _row("2025-06-01", "5000", "Materials", "ACME LTD", debit=500.0, reference="INV-777"),
    ]
    nominal = pd.DataFrame(rows)
    result = ad.duplicate_transactions(nominal)

    assert result.status == "review"
    assert len(result.detail) == 2
    assert all("reference" in r.lower() for r in result.detail["Flag reason"])


def test_duplicate_transactions_ignores_double_entry_legs_of_one_invoice():
    # the same reference posted to two different codes with opposite signs
    # (Sales credit + Debtors debit) is just double-entry, not a duplicate;
    # neither is an invoice and its later payment sharing a reference on
    # the same control account with opposite signs
    rows = [
        _row("2025-02-05", "0010", "Sales", "GARDEN CENTRE LTD", credit=12400.0, reference="1001/2025", description="Sales invoice"),
        _row("2025-02-05", "610A", "Accounts Receivable", "GARDEN CENTRE LTD", debit=12400.0, reference="1001/2025", description="Sales invoice"),
        _row("2025-03-01", "610A", "Accounts Receivable", "GARDEN CENTRE LTD", credit=12400.0, reference="1001/2025", description="Payment received"),
    ]
    nominal = pd.DataFrame(rows)
    result = ad.duplicate_transactions(nominal)
    assert result.status == "ok"
    assert result.detail.empty


def test_duplicate_transactions_ok_when_nothing_repeats():
    rows = [
        _row("2025-03-01", "5000", "Materials", "ACME LTD", debit=500.0, reference="INV-001"),
        _row("2025-04-01", "5000", "Materials", "ACME LTD", debit=300.0, reference="INV-002"),
    ]
    nominal = pd.DataFrame(rows)
    result = ad.duplicate_transactions(nominal)
    assert result.status == "ok"
    assert result.detail.empty


def test_unusual_journal_posting_dates_flags_weekend_journals_only():
    rows = [
        _row("2025-03-01", "7000", "Payroll", "N/A", debit=1000.0, source_type="Journal"),   # Saturday
        _row("2025-03-03", "7000", "Payroll", "N/A", debit=1000.0, source_type="Journal"),   # Monday - fine
        _row("2025-03-02", "5000", "Materials", "ACME LTD", debit=500.0, source_type="Bill"),  # Sunday but not a journal
    ]
    nominal = pd.DataFrame(rows)
    result = ad.unusual_journal_posting_dates(nominal)

    assert result.status == "review"
    assert len(result.detail) == 1
    assert result.detail.iloc[0]["Date"] == pd.Timestamp("2025-03-01")


def test_unusual_journal_posting_dates_ok_when_none_on_weekend():
    rows = [_row("2025-03-03", "7000", "Payroll", "N/A", debit=1000.0, source_type="Journal")]
    nominal = pd.DataFrame(rows)
    result = ad.unusual_journal_posting_dates(nominal)
    assert result.status == "ok"


def test_run_all_anomaly_checks_returns_three_results_with_no_data():
    results = ad.run_all_anomaly_checks(pd.DataFrame())
    assert len(results) == 3
    assert all(r.status == "n/a" for r in results)
