"""Unit tests for the Related Party Transactions check - see
app/related_party_transactions.py: identifies related parties from
Directors' Loan Account contact names (this system's only source of
related-party identity), then scans the rest of nominal activity for
postings to those same contacts outside the DLA account."""
import pandas as pd

from app.related_party_transactions import find_related_party_transactions


def _tb(rows):
    return pd.DataFrame(rows)


def _nom(date, code, name, debit=0.0, credit=0.0, description="", contact=""):
    return {"date": pd.Timestamp(date), "account_code": code, "account_name": name,
            "reference": "", "description": description, "contact": contact, "debit": debit, "credit": credit}


def test_flags_a_posting_to_a_dla_contact_outside_the_dla_account():
    tb_current = _tb([{"account_code": "8100", "account_name": "DIRECTORS LOAN ACCOUNT", "account_type": "Current Liability"}])
    nominal = pd.DataFrame([
        _nom("2025-03-01", "8100", "DIRECTORS LOAN ACCOUNT", credit=3000.0, contact="J Smith"),
        _nom("2025-06-01", "7200", "RENT", debit=12000.0, contact="J Smith", description="Rent to director"),
    ])
    result = find_related_party_transactions(tb_current, nominal)
    assert result.status == "review"
    assert len(result.detail) == 1
    assert result.detail.iloc[0]["Contact"] == "J Smith"
    assert "FRS 102 Section 33" in result.message
    assert "Not a completeness check" in result.message


def test_deduplicates_the_double_entry_contra_leg_of_the_same_posting():
    tb_current = _tb([{"account_code": "8100", "account_name": "DIRECTORS LOAN ACCOUNT", "account_type": "Current Liability"}])
    nominal = pd.DataFrame([
        _nom("2025-03-01", "8100", "DIRECTORS LOAN ACCOUNT", credit=3000.0, contact="J Smith"),
        _nom("2025-06-01", "7200", "RENT", debit=12000.0, contact="J Smith", description="Rent to director"),
        _nom("2025-06-01", "1200", "BANK", credit=12000.0, contact="J Smith", description="Rent to director"),
    ])
    result = find_related_party_transactions(tb_current, nominal)
    assert result.status == "review"
    assert len(result.detail) == 1  # not 2 - same real transaction, two double-entry legs


def test_ignores_unrelated_contacts():
    tb_current = _tb([{"account_code": "8100", "account_name": "DIRECTORS LOAN ACCOUNT", "account_type": "Current Liability"}])
    nominal = pd.DataFrame([
        _nom("2025-03-01", "8100", "DIRECTORS LOAN ACCOUNT", credit=3000.0, contact="J Smith"),
        _nom("2025-07-01", "7300", "OFFICE COSTS", debit=600.0, contact="Staples"),
    ])
    result = find_related_party_transactions(tb_current, nominal)
    assert result.status == "ok"


def test_ignores_postings_below_threshold():
    tb_current = _tb([{"account_code": "8100", "account_name": "DIRECTORS LOAN ACCOUNT", "account_type": "Current Liability"}])
    nominal = pd.DataFrame([
        _nom("2025-03-01", "8100", "DIRECTORS LOAN ACCOUNT", credit=3000.0, contact="J Smith"),
        _nom("2025-06-01", "7300", "OFFICE COSTS", debit=50.0, contact="J Smith"),
    ])
    result = find_related_party_transactions(tb_current, nominal, threshold=500.0)
    assert result.status == "ok"


def test_na_without_a_dla_account():
    tb_current = _tb([{"account_code": "1", "account_name": "SALES", "account_type": "Sales"}])
    nominal = pd.DataFrame([_nom("2025-01-01", "1", "SALES", credit=1000.0, contact="Acme Ltd")])
    result = find_related_party_transactions(tb_current, nominal)
    assert result.status == "n/a"


def test_na_without_named_dla_contacts():
    tb_current = _tb([{"account_code": "8100", "account_name": "DIRECTORS LOAN ACCOUNT", "account_type": "Current Liability"}])
    nominal = pd.DataFrame([_nom("2025-03-01", "8100", "DIRECTORS LOAN ACCOUNT", credit=3000.0, contact="")])
    result = find_related_party_transactions(tb_current, nominal)
    assert result.status == "n/a"
