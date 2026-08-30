"""Unit tests for the fixed asset register robustness features added on
top of category_level_rollforward - see app/fixed_assets.py:

  - far_additions_detail: the actual transaction-level postings behind a
    category's "Additions" total, so a preparer can check the figure
    against real transactions rather than just trust an aggregate.
  - the "System est." columns on category_level_rollforward: a default
    depreciation rate/method inferred from the category name, shown as
    an advisory sanity-check against what's actually booked - never
    itself a reason to flag a category "review".
  - suggest_capital_expenditure_reclassification: scans expense-coded
    nominal activity for postings that look like capital expenditure
    miscoded to a P&L account, using a vocabulary built from this
    client's own fixed asset categories plus common capex nouns.

Existing category_level_rollforward/asset_level_rollforward behaviour
(grouping, TB-tie checks, per-asset depreciation) is covered by
tests/test_pipeline.py and tests/test_formulas.py - not repeated here."""
import pandas as pd
import pytest

from app import fixed_assets as fa


def _fa_tb_row(code, name, account_type, debit, credit):
    return {"account_code": code, "account_name": name, "account_type": account_type,
            "debit": debit, "credit": credit, "balance": debit - credit}


def _nom_row(date, code, name, debit=0.0, credit=0.0, description="", reference="", contact=""):
    return {"date": pd.Timestamp(date), "account_code": code, "account_name": name,
            "reference": reference, "description": description, "contact": contact,
            "debit": debit, "credit": credit}


# --- far_additions_detail / category rollforward extra_detail -----------

def test_far_additions_detail_lists_transactions_behind_the_additions_total():
    tb_current = pd.DataFrame([
        _fa_tb_row("6350", "COMPUTER EQUIPMENT - COST", "Fixed Asset", 5000.0, 0.0),
        _fa_tb_row("6360", "COMPUTER EQUIPMENT - DEPRECIATION", "Fixed Asset", 0.0, 1000.0),
    ])
    nominal = pd.DataFrame([
        _nom_row("2025-06-01", "6350", "COMPUTER EQUIPMENT - COST", debit=1200.0, contact="PC World", description="Laptops"),
        _nom_row("2025-07-01", "6360", "COMPUTER EQUIPMENT - DEPRECIATION", credit=200.0),  # depreciation charge, not an addition
    ])
    detail = fa.far_additions_detail(tb_current, nominal)
    assert len(detail) == 1
    row = detail.iloc[0]
    assert row["Category"] == "COMPUTER EQUIPMENT"
    assert row["Addition (Cost)"] == 1200.0
    assert row["Contact"] == "PC World"


def test_far_additions_detail_empty_when_no_fixed_asset_accounts():
    assert fa.far_additions_detail(pd.DataFrame(), pd.DataFrame()).empty
    tb_current = pd.DataFrame([_fa_tb_row("4000", "SALES", "Sales", 0.0, 1000.0)])
    nominal = pd.DataFrame([_nom_row("2025-06-01", "4000", "SALES", credit=1000.0)])
    assert fa.far_additions_detail(tb_current, nominal).empty


def test_category_rollforward_attaches_additions_as_extra_detail():
    tb_current = pd.DataFrame([_fa_tb_row("6350", "COMPUTER EQUIPMENT - COST", "Fixed Asset", 1200.0, 0.0)])
    nominal = pd.DataFrame([_nom_row("2025-06-01", "6350", "COMPUTER EQUIPMENT - COST", debit=1200.0, contact="PC World")])
    result = fa.category_level_rollforward(tb_current, pd.DataFrame(), nominal)
    assert not result.extra_detail.empty
    assert "additions posted during the year" in result.extra_detail_label.lower()
    assert result.extra_detail.iloc[0]["Contact"] == "PC World"


def test_category_rollforward_has_no_extra_detail_when_no_additions():
    tb_current = pd.DataFrame([_fa_tb_row("6350", "COMPUTER EQUIPMENT - COST", "Fixed Asset", 0.0, 0.0)])
    result = fa.category_level_rollforward(tb_current, pd.DataFrame(), pd.DataFrame())
    assert result.extra_detail.empty


# --- System-estimated depreciation ---------------------------------------

def test_infer_category_rate_matches_specific_before_generic_keyword():
    # "computer equipment" should match the specific "computer" rate, not
    # fall through to the generic "equipment" rate further down the dict
    method, rate = fa._infer_category_rate("COMPUTER EQUIPMENT")
    assert (method, rate) == (fa.REDUCING_BALANCE, 25.0)


def test_infer_category_rate_falls_back_for_unrecognised_category():
    assert fa._infer_category_rate("SOME UNUSUAL CATEGORY") == fa._DEFAULT_RATE_FALLBACK


def test_category_rollforward_system_estimate_columns_present_and_advisory_only():
    # motor vehicles -> reducing balance 25% by default; deliberately
    # booked with NO depreciation charge at all, to prove the mismatch is
    # surfaced as a number/advisory note, never as a "review" status by
    # itself (the category still ties to the TB, so it should stay "ok")
    tb_current = pd.DataFrame([
        _fa_tb_row("7100", "MOTOR VEHICLES - COST", "Fixed Asset", 18000.0, 0.0),
        _fa_tb_row("7110", "MOTOR VEHICLES - DEPRECIATION", "Fixed Asset", 0.0, 0.0),
    ])
    tb_comparative = pd.DataFrame([
        _fa_tb_row("7100", "MOTOR VEHICLES - COST", "Fixed Asset", 18000.0, 0.0),
        _fa_tb_row("7110", "MOTOR VEHICLES - DEPRECIATION", "Fixed Asset", 0.0, 0.0),
    ])
    result = fa.category_level_rollforward(tb_current, tb_comparative, pd.DataFrame(), period_days=365)
    row = result.detail.iloc[0]
    assert row["System est. method"] == "Reducing balance"
    assert row["System est. rate %"] == 25.0
    assert row["System est. depreciation"] == pytest.approx(18000.0 * 0.25)
    assert row["Booked vs system est. (diff)"] == pytest.approx(0.0 - 18000.0 * 0.25)
    assert result.status == "ok"  # no TB-tie discrepancy - the estimate mismatch alone never flips this
    assert "system est" in result.message.lower()


def test_system_estimate_prorates_for_a_short_period():
    tb_current = pd.DataFrame([_fa_tb_row("7100", "MOTOR VEHICLES - COST", "Fixed Asset", 18000.0, 0.0)])
    tb_comparative = pd.DataFrame([_fa_tb_row("7100", "MOTOR VEHICLES - COST", "Fixed Asset", 18000.0, 0.0)])
    full_year = fa.category_level_rollforward(tb_current, tb_comparative, pd.DataFrame(), period_days=365)
    half_year = fa.category_level_rollforward(tb_current, tb_comparative, pd.DataFrame(), period_days=182)
    full_est = full_year.detail.iloc[0]["System est. depreciation"]
    half_est = half_year.detail.iloc[0]["System est. depreciation"]
    assert half_est == pytest.approx(full_est * (182 / 365), abs=0.5)


# --- suggest_capital_expenditure_reclassification ------------------------

def test_capex_suggestion_flags_expense_posting_matching_client_own_category():
    tb_current = pd.DataFrame([
        _fa_tb_row("6350", "COMPUTER EQUIPMENT - COST", "Fixed Asset", 5000.0, 0.0),
        _fa_tb_row("7100", "IT COSTS", "Overhead", 0.0, 0.0),
    ])
    nominal = pd.DataFrame([
        _nom_row("2025-07-15", "7100", "IT COSTS", debit=899.0, description="Dell laptop for new starter", contact="Dell"),
    ])
    result = fa.suggest_capital_expenditure_reclassification(tb_current, nominal)
    assert result.status == "review"
    assert len(result.detail) == 1
    assert "laptop" in result.detail.iloc[0]["Matched on"]


def test_capex_suggestion_ignores_postings_below_threshold():
    tb_current = pd.DataFrame([
        _fa_tb_row("6350", "COMPUTER EQUIPMENT - COST", "Fixed Asset", 5000.0, 0.0),
        _fa_tb_row("7100", "IT COSTS", "Overhead", 0.0, 0.0),
    ])
    nominal = pd.DataFrame([
        _nom_row("2025-07-15", "7100", "IT COSTS", debit=45.0, description="Laptop stand", contact="Amazon"),
    ])
    result = fa.suggest_capital_expenditure_reclassification(tb_current, nominal, threshold=500.0)
    assert result.status == "ok"


def test_capex_suggestion_ignores_postings_already_coded_to_fixed_asset():
    tb_current = pd.DataFrame([_fa_tb_row("6350", "COMPUTER EQUIPMENT - COST", "Fixed Asset", 5000.0, 0.0)])
    nominal = pd.DataFrame([
        _nom_row("2025-07-15", "6350", "COMPUTER EQUIPMENT - COST", debit=899.0, description="New laptop"),
    ])
    result = fa.suggest_capital_expenditure_reclassification(tb_current, nominal)
    assert result.status == "ok"  # already in the register, not "coded elsewhere"


def test_capex_suggestion_ignores_non_expense_account_types():
    # a big balance-sheet posting (e.g. a director's loan repayment)
    # mentioning "van" shouldn't be flagged - it isn't coded to an
    # expense-like account at all
    tb_current = pd.DataFrame([
        _fa_tb_row("6350", "MOTOR VEHICLES - COST", "Fixed Asset", 5000.0, 0.0),
        _fa_tb_row("2100", "DIRECTORS LOAN ACCOUNT", "Current Liability", 0.0, 0.0),
    ])
    nominal = pd.DataFrame([
        _nom_row("2025-07-15", "2100", "DIRECTORS LOAN ACCOUNT", debit=15000.0, description="Repayment re: van purchase"),
    ])
    result = fa.suggest_capital_expenditure_reclassification(tb_current, nominal)
    assert result.status == "ok"


def test_capex_suggestion_uses_generic_vocabulary_when_client_has_no_fa_categories_yet():
    # a brand-new client with nothing coded to Fixed Asset at all yet -
    # the generic capex noun list is still enough to catch an obvious one
    tb_current = pd.DataFrame([_fa_tb_row("7100", "REPAIRS AND MAINTENANCE", "Overhead", 0.0, 0.0)])
    nominal = pd.DataFrame([
        _nom_row("2025-07-15", "7100", "REPAIRS AND MAINTENANCE", debit=12000.0, description="New delivery van"),
    ])
    result = fa.suggest_capital_expenditure_reclassification(tb_current, nominal)
    assert result.status == "review"
    assert "van" in result.detail.iloc[0]["Matched on"]


def test_capex_suggestion_uses_prior_register_categories_too():
    tb_current = pd.DataFrame([_fa_tb_row("7100", "SUNDRY EXPENSES", "Overhead", 0.0, 0.0)])
    nominal = pd.DataFrame([
        _nom_row("2025-07-15", "7100", "SUNDRY EXPENSES", debit=2500.0, description="New forklift for warehouse"),
    ])
    prior_register = pd.DataFrame([{"asset_id": "FA-1", "category": "Warehouse Forklifts", "cost": 20000.0}])
    result = fa.suggest_capital_expenditure_reclassification(tb_current, nominal, prior_register=prior_register)
    assert result.status == "review"
    assert "forklift" in result.detail.iloc[0]["Matched on"]


def test_capex_suggestion_na_without_tb_or_nominal_activity():
    assert fa.suggest_capital_expenditure_reclassification(None, None).status == "n/a"
    assert fa.suggest_capital_expenditure_reclassification(pd.DataFrame(), pd.DataFrame()).status == "n/a"
