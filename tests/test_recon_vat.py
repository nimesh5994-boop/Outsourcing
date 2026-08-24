"""Unit tests for the VAT cross-check's candidate-reconciling-items
enhancement: when the VAT return doesn't tie to the nominal ledger within
materiality, the flagged check should come with the actual nominal
activity postings to VAT-related accounts attached, not just a generic
"check for timing differences" message with nothing to check against."""
import pandas as pd

from app.recon import vat_cross_check

VAT_CONTROL_NAMES = ["vat control", "vat liability"]


def _vat_return(box5=800.0, box6=15000.0):
    return pd.DataFrame([{
        "box1": 1000.0, "box2": 0.0, "box3": 1000.0, "box4": 200.0,
        "box5": box5, "box6": box6, "box7": 5000.0, "box8": 0.0, "box9": 0.0,
    }])


def _tb(vat_control_balance=-500.0):
    return pd.DataFrame([
        {"account_code": "2200", "account_name": "VAT Control", "account_type": "Current Liability",
         "debit": 0.0, "credit": abs(vat_control_balance), "balance": vat_control_balance},
    ])


def _pl(turnover=15000.0):
    return pd.DataFrame([{"account_code": "4000", "account_name": "Sales", "category": "Turnover", "amount": turnover}])


def _nominal(rows):
    df = pd.DataFrame(rows)
    for col in ("date", "account_code", "account_name", "reference", "description", "contact", "source_type", "debit", "credit"):
        if col not in df.columns:
            df[col] = "" if col not in ("debit", "credit") else 0.0
    return df


def test_ok_status_has_no_extra_detail():
    # box5 (800) ties to |vat control balance| (800) within materiality
    result = vat_cross_check(_vat_return(box5=800.0), _pl(), _tb(vat_control_balance=-800.0), VAT_CONTROL_NAMES)
    assert result.status == "ok"
    assert result.extra_detail.empty


def test_flagged_variance_surfaces_candidate_nominal_postings():
    nominal = _nominal([
        {"date": pd.Timestamp("2025-03-01"), "account_name": "VAT Control", "description": "Qtr VAT journal", "contact": "", "debit": 0.0, "credit": 700.0},
        {"date": pd.Timestamp("2025-03-15"), "account_name": "Sales", "description": "Invoice 101", "contact": "Acme Ltd", "debit": 0.0, "credit": 1200.0},
    ])
    # box5 (800) vs |vat control balance| (700) -> £100 variance, below materiality (£500) so use a bigger gap
    result = vat_cross_check(_vat_return(box5=1500.0), _pl(), _tb(vat_control_balance=-700.0), VAT_CONTROL_NAMES, nominal)
    assert result.status == "review"
    assert not result.extra_detail.empty
    # only the VAT-related posting should be included, not the unrelated Sales one
    assert list(result.extra_detail["account_name"]) == ["VAT Control"]
    assert result.extra_detail_label  # non-empty, describes what the table is


def test_flagged_variance_with_no_nominal_activity_has_no_extra_detail():
    result = vat_cross_check(_vat_return(box5=1500.0), _pl(), _tb(vat_control_balance=-700.0), VAT_CONTROL_NAMES, None)
    assert result.status == "review"
    assert result.extra_detail.empty


def test_flagged_variance_with_no_matching_postings_has_no_extra_detail():
    nominal = _nominal([
        {"date": pd.Timestamp("2025-03-15"), "account_name": "Sales", "description": "Invoice 101", "contact": "Acme Ltd", "debit": 0.0, "credit": 1200.0},
    ])
    result = vat_cross_check(_vat_return(box5=1500.0), _pl(), _tb(vat_control_balance=-700.0), VAT_CONTROL_NAMES, nominal)
    assert result.status == "review"
    assert result.extra_detail.empty
