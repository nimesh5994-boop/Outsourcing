"""Unit tests for the control account rollforward robustness features
added on top of build_rollforward - see app/control_accounts.py:

  - extra_detail: the actual nominal-activity postings behind the
    schedule's single "MOVEMENTS DURING YEAR" total, so a preparer can
    check real transactions rather than just trust an aggregate.
  - movement_breakdown: a net-movement-by-contact view for control
    accounts that have no aged debtors/creditors listing to break their
    closing balance down against - honestly labelled as the year's
    movement only, never checked against the TB balance the way the
    aged-listing breakdown is.
  - suggest_control_account_miscoding: scans nominal activity for a
    contact known to one control account posting, above materiality, to
    a DIFFERENT balance-sheet control-account-shaped account - never
    Bank or a P&L account, since those are the normal contra-leg of a
    correctly-coded transaction and would just be noise.

Existing build_rollforward/build_all_rollforwards behaviour (TB-tie
checks, the debtors/creditors aged-listing breakdown) is covered by
tests/test_pipeline.py and tests/test_formulas.py - not repeated here."""
import pandas as pd

from app import control_accounts as ca


def _nom_row(date, code, name, debit=0.0, credit=0.0, description="", reference="", contact=""):
    return {"date": pd.Timestamp(date), "account_code": code, "account_name": name,
            "reference": reference, "description": description, "contact": contact,
            "debit": debit, "credit": credit}


# --- find_control_accounts account-type coverage --------------------------

def test_find_control_accounts_covers_prepayment_inventory_and_non_current_types():
    # a regression test for a real gap: Prepayment, Inventory, Non-current
    # Asset and Non-current Liability are real, standard Xero account
    # types that were previously entirely missing from control-account
    # discovery - not because a prepayment or stock account isn't a
    # control account, but because its type string just wasn't in the set
    tb_current = pd.DataFrame([
        {"account_code": "620", "account_name": "PREPAYMENTS", "account_type": "Prepayment"},
        {"account_code": "630", "account_name": "STOCK ON HAND", "account_type": "Inventory"},
        {"account_code": "900", "account_name": "BANK LOAN", "account_type": "Non-current Liability"},
        {"account_code": "910", "account_name": "LONG TERM DEPOSIT", "account_type": "Non-current Asset"},
        {"account_code": "6350", "account_name": "COMPUTER EQUIPMENT - COST", "account_type": "Fixed Asset"},
    ])
    nominal = pd.DataFrame([
        _nom_row("2025-06-01", "620", "PREPAYMENTS", debit=2400.0, contact="AXA Insurance"),
        _nom_row("2025-07-01", "630", "STOCK ON HAND", debit=5000.0, contact="Supplier Ltd"),
        _nom_row("2025-08-01", "900", "BANK LOAN", credit=1000.0, contact="Barclays"),
        _nom_row("2025-09-01", "910", "LONG TERM DEPOSIT", debit=1000.0, contact="Landlord"),
        _nom_row("2025-10-01", "6350", "COMPUTER EQUIPMENT - COST", debit=1200.0, contact="Dell"),
    ])
    codes = {code for code, _ in ca.find_control_accounts(tb_current, nominal)}
    assert codes == {"620", "630", "900", "910"}  # fixed assets excluded - handled by fixed_assets.py instead


# --- extra_detail / movement_breakdown -----------------------------------

def test_build_rollforward_attaches_movement_transactions_as_extra_detail():
    tb_current = pd.DataFrame([{"account_code": "8100", "account_name": "DIRECTORS LOAN ACCOUNT", "account_type": "Current Liability", "balance": -3500.0}])
    tb_comparative = pd.DataFrame([{"account_code": "8100", "account_name": "DIRECTORS LOAN ACCOUNT", "account_type": "Current Liability", "balance": 0.0}])
    nominal = pd.DataFrame([
        _nom_row("2025-03-01", "8100", "DIRECTORS LOAN ACCOUNT", credit=3000.0, contact="J Smith", description="Drawdown"),
        _nom_row("2025-06-01", "8100", "DIRECTORS LOAN ACCOUNT", credit=500.0, contact="J Smith", description="Drawdown"),
    ])
    result = ca.build_rollforward("8100", "DIRECTORS LOAN ACCOUNT", tb_current, tb_comparative, nominal)
    assert len(result.extra_detail) == 2
    assert set(result.extra_detail["Contact"]) == {"J Smith"}
    assert "postings behind the year's movement" in result.extra_detail_label.lower()


def test_movement_breakdown_populated_only_when_no_aged_listing_breakdown():
    tb_current = pd.DataFrame([{"account_code": "8100", "account_name": "DIRECTORS LOAN ACCOUNT", "account_type": "Current Liability", "balance": -3000.0}])
    tb_comparative = pd.DataFrame()
    nominal = pd.DataFrame([_nom_row("2025-03-01", "8100", "DIRECTORS LOAN ACCOUNT", credit=3000.0, contact="J Smith")])
    result = ca.build_rollforward("8100", "DIRECTORS LOAN ACCOUNT", tb_current, tb_comparative, nominal)
    assert not result.movement_breakdown.empty
    assert result.movement_breakdown.iloc[0]["Party"] == "J Smith"
    assert result.movement_breakdown.iloc[0]["Net movement £"] == -3000.0
    assert result.breakdown.empty
    assert "not the closing balance make-up" in result.movement_breakdown_label.lower()


def test_movement_breakdown_absent_for_debtors_control_which_has_its_own_breakdown():
    tb_current = pd.DataFrame([{"account_code": "1100", "account_name": "TRADE DEBTORS CONTROL", "account_type": "Current Asset", "balance": 2370.0}])
    tb_comparative = pd.DataFrame()
    nominal = pd.DataFrame([_nom_row("2025-03-01", "1100", "TRADE DEBTORS CONTROL", debit=2370.0, contact="Moss & Turf Supplies")])
    aged_debtors = pd.DataFrame([{"customer": "MOSS & TURF SUPPLIES", "total": 2370.0}])
    result = ca.build_rollforward("1100", "TRADE DEBTORS CONTROL", tb_current, tb_comparative, nominal, aged_debtors=aged_debtors)
    assert not result.breakdown.empty  # the authoritative aged-listing breakdown
    assert result.movement_breakdown.empty  # no second, differently-scoped one alongside it


def test_movement_breakdown_empty_when_no_movement():
    tb_current = pd.DataFrame([{"account_code": "8100", "account_name": "DIRECTORS LOAN ACCOUNT", "account_type": "Current Liability", "balance": 0.0}])
    result = ca.build_rollforward("8100", "DIRECTORS LOAN ACCOUNT", tb_current, pd.DataFrame(), pd.DataFrame())
    assert result.extra_detail.empty
    assert result.movement_breakdown.empty


# --- suggest_control_account_miscoding ------------------------------------

def _tb(rows):
    return pd.DataFrame(rows)


def test_miscoding_suggestion_flags_posting_to_a_different_control_account():
    tb_current = _tb([
        {"account_code": "8100", "account_name": "DIRECTORS LOAN ACCOUNT", "account_type": "Current Liability", "balance": -3000.0},
        {"account_code": "8200", "account_name": "OTHER CREDITORS", "account_type": "Current Liability", "balance": -1200.0},
        {"account_code": "1200", "account_name": "BANK", "account_type": "Bank", "balance": 20000.0},
    ])
    nominal = pd.DataFrame([
        _nom_row("2025-03-01", "8100", "DIRECTORS LOAN ACCOUNT", credit=3000.0, contact="J Smith", description="Drawdown"),
        _nom_row("2025-03-01", "1200", "BANK", debit=3000.0, contact="J Smith", description="Drawdown"),
        _nom_row("2025-08-01", "8200", "OTHER CREDITORS", credit=1200.0, contact="J Smith", description="Repayment to J Smith"),
    ])
    control_accounts = [("8100", "DIRECTORS LOAN ACCOUNT"), ("8200", "OTHER CREDITORS")]
    result = ca.suggest_control_account_miscoding(tb_current, nominal, control_accounts)
    assert result.status == "review"
    assert len(result.detail) == 1
    row = result.detail.iloc[0]
    assert row["Nominal Code"] == "8200"
    assert row["Normally clears through"] == "DIRECTORS LOAN ACCOUNT"


def test_miscoding_suggestion_ignores_normal_bank_and_pl_contra_legs():
    # the Bank receipt and the Sales invoice are the NORMAL double-entry
    # contra-leg of correctly-coded transactions - neither should ever be
    # flagged, or every genuine posting would trigger a false positive
    tb_current = _tb([
        {"account_code": "1100", "account_name": "TRADE DEBTORS CONTROL", "account_type": "Current Asset", "balance": 0.0},
        {"account_code": "1200", "account_name": "BANK", "account_type": "Bank", "balance": 5000.0},
        {"account_code": "4000", "account_name": "SALES", "account_type": "Sales", "balance": -5000.0},
    ])
    nominal = pd.DataFrame([
        _nom_row("2025-04-01", "1100", "TRADE DEBTORS CONTROL", debit=5000.0, contact="Acme Ltd", description="Invoice"),
        _nom_row("2025-04-01", "4000", "SALES", credit=5000.0, contact="Acme Ltd", description="Invoice"),
        _nom_row("2025-05-01", "1200", "BANK", debit=5000.0, contact="Acme Ltd", description="Receipt"),
        _nom_row("2025-05-01", "1100", "TRADE DEBTORS CONTROL", credit=5000.0, contact="Acme Ltd", description="Receipt"),
    ])
    control_accounts = [("1100", "TRADE DEBTORS CONTROL")]
    result = ca.suggest_control_account_miscoding(tb_current, nominal, control_accounts)
    assert result.status == "ok"


def test_miscoding_suggestion_ignores_postings_below_threshold():
    tb_current = _tb([
        {"account_code": "8100", "account_name": "DIRECTORS LOAN ACCOUNT", "account_type": "Current Liability", "balance": -3000.0},
        {"account_code": "8200", "account_name": "OTHER CREDITORS", "account_type": "Current Liability", "balance": -50.0},
    ])
    nominal = pd.DataFrame([
        _nom_row("2025-03-01", "8100", "DIRECTORS LOAN ACCOUNT", credit=3000.0, contact="J Smith"),
        _nom_row("2025-08-01", "8200", "OTHER CREDITORS", credit=50.0, contact="J Smith"),
    ])
    control_accounts = [("8100", "DIRECTORS LOAN ACCOUNT"), ("8200", "OTHER CREDITORS")]
    result = ca.suggest_control_account_miscoding(tb_current, nominal, control_accounts, threshold=500.0)
    assert result.status == "ok"


def test_miscoding_suggestion_uses_aged_listing_as_independent_ground_truth():
    # the aged debtors listing names Acme Ltd as a customer even though
    # they have no nominal activity in Trade Debtors Control this year -
    # a payment received from them coded to Other Debtors should still
    # be caught, because the aged listing is independent ground truth
    tb_current = _tb([
        {"account_code": "1100", "account_name": "TRADE DEBTORS CONTROL", "account_type": "Current Asset", "balance": 0.0},
        {"account_code": "1150", "account_name": "OTHER DEBTORS", "account_type": "Current Asset", "balance": 900.0},
    ])
    nominal = pd.DataFrame([
        _nom_row("2025-06-01", "1150", "OTHER DEBTORS", debit=900.0, contact="Acme Ltd", description="Miscellaneous receivable"),
    ])
    aged_debtors = pd.DataFrame([{"customer": "Acme Ltd", "total": 0.0}])
    control_accounts = [("1100", "TRADE DEBTORS CONTROL"), ("1150", "OTHER DEBTORS")]
    result = ca.suggest_control_account_miscoding(tb_current, nominal, control_accounts, aged_debtors=aged_debtors)
    assert result.status == "review"
    assert result.detail.iloc[0]["Normally clears through"] == "TRADE DEBTORS CONTROL"


def test_miscoding_suggestion_na_without_required_inputs():
    assert ca.suggest_control_account_miscoding(None, None, []).status == "n/a"
    assert ca.suggest_control_account_miscoding(pd.DataFrame(), pd.DataFrame(), []).status == "n/a"
    tb_current = _tb([{"account_code": "8100", "account_name": "DIRECTORS LOAN ACCOUNT", "account_type": "Current Liability", "balance": 0.0}])
    nominal = pd.DataFrame([_nom_row("2025-03-01", "8100", "DIRECTORS LOAN ACCOUNT", credit=100.0, contact="J Smith")])
    assert ca.suggest_control_account_miscoding(tb_current, nominal, []).status == "n/a"
