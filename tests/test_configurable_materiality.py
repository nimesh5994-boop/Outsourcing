"""Unit tests proving materiality is genuinely configurable per-template
rather than a fixed module constant - see storage.DEFAULT_TEMPLATE_CONFIG's
"materiality" block and main.py's _generate_workbook_steps, which reads it
and threads a `materiality`/`variance_pct_threshold` value through every
check module. Each test below picks a variance that sits BELOW the
module's own default MATERIALITY_AMOUNT (so it's "ok" out of the box) and
proves a tighter, explicitly-passed materiality flips it to "review" -
the concrete evidence that the parameter is actually read, not just
accepted and ignored."""
import pandas as pd

from app.accruals_prepayments import build_schedule as accruals_build_schedule
from app.control_accounts import build_all_rollforwards, build_rollforward
from app.corporation_tax import compute as ct_compute
from app.financial_statements import build_bs_statement
from app.fixed_assets import category_level_rollforward
from app.going_concern import assess as going_concern_assess
from app.nominal_matrix import build_matrix
from app.recon import bank_reconciliation, debtors_creditors_control_recon, run_all_recons, variance_analysis
from app.related_party_transactions import find_related_party_transactions


def _tb(rows):
    return pd.DataFrame(rows)


def _nom(date, code, name, debit=0.0, credit=0.0, description="", contact=""):
    return {"date": pd.Timestamp(date), "account_code": code, "account_name": name,
            "reference": "", "description": description, "contact": contact, "debit": debit, "credit": credit}


def test_debtors_control_recon_respects_a_tighter_materiality():
    aged = pd.DataFrame([{"customer": "Acme Ltd", "current": 800.0, "bucket_1": 0.0, "bucket_2": 0.0, "bucket_3": 0.0, "bucket_4": 0.0, "older": 0.0, "total": 800.0}])
    tb = pd.DataFrame([{"account_code": "1100", "account_name": "TRADE DEBTORS CONTROL", "account_type": "Current Asset", "balance": 900.0}])
    # variance is £100 - "ok" at the default £500 materiality, "review" once
    # a practice tightens it to £50
    default_result = debtors_creditors_control_recon(aged, tb, ["debtors control", "trade debtors"], "customer", "Debtors recon")
    assert default_result.status == "ok"
    tight_result = debtors_creditors_control_recon(aged, tb, ["debtors control", "trade debtors"], "customer", "Debtors recon", materiality=50.0)
    assert tight_result.status == "review"


def test_variance_analysis_respects_configurable_materiality_and_pct_threshold():
    tb_current = _tb([{"account_code": "7000", "account_name": "RENT", "balance": 1100.0}])
    tb_comparative = _tb([{"account_code": "7000", "account_name": "RENT", "balance": 1000.0}])
    # £100 / 10% movement - below the default £500 absolute threshold
    default_result = variance_analysis(tb_current, tb_comparative)
    assert default_result.status == "ok"
    tight_result = variance_analysis(tb_current, tb_comparative, materiality=50.0, variance_pct_threshold=0.05)
    assert tight_result.status == "review"


def test_bank_reconciliation_respects_configurable_materiality():
    bank_statement = pd.DataFrame([{"account_name": "Bank Current Account", "statement_date": "2025-12-31", "closing_balance": 10100.0}])
    tb = pd.DataFrame([{"account_code": "1200", "account_name": "Bank Current Account", "balance": 10000.0}])
    assert bank_reconciliation(bank_statement, tb).status == "ok"
    assert bank_reconciliation(bank_statement, tb, materiality=50.0).status == "review"


def test_run_all_recons_threads_materiality_into_every_sub_check():
    data = {
        "tb_current": _tb([{"account_code": "7000", "account_name": "RENT", "balance": 1100.0, "debit": 1100.0, "credit": 0.0}]),
        "tb_comparative": _tb([{"account_code": "7000", "account_name": "RENT", "balance": 1000.0, "debit": 1000.0, "credit": 0.0}]),
    }
    default_results = run_all_recons(data)
    variance = next(r for r in default_results if r.name == "Current vs comparative variance analysis")
    assert variance.status == "ok"
    tight_results = run_all_recons(data, materiality=50.0, variance_pct_threshold=0.05)
    variance_tight = next(r for r in tight_results if r.name == "Current vs comparative variance analysis")
    assert variance_tight.status == "review"


def test_control_account_rollforward_ties_within_configurable_materiality():
    tb_comparative = _tb([{"account_code": "8100", "account_name": "DIRECTORS LOAN ACCOUNT", "account_type": "Current Liability", "balance": -1000.0}])
    tb_current = _tb([{"account_code": "8100", "account_name": "DIRECTORS LOAN ACCOUNT", "account_type": "Current Liability", "balance": -1150.0}])
    nominal = pd.DataFrame([_nom("2025-06-01", "8100", "DIRECTORS LOAN ACCOUNT", credit=100.0, contact="J Smith")])
    # b/fwd -1000 + movement -100 = -1100 computed vs -1150 per TB: £50 gap
    default_result = build_rollforward("8100", "DIRECTORS LOAN ACCOUNT", tb_current, tb_comparative, nominal)
    assert default_result.status == "ok"
    tight_result = build_rollforward("8100", "DIRECTORS LOAN ACCOUNT", tb_current, tb_comparative, nominal, materiality=10.0)
    assert tight_result.status == "review"

    # build_all_rollforwards threads the same param through to build_rollforward
    all_default = build_all_rollforwards(tb_current, tb_comparative, nominal)
    assert all_default[0].status == "ok"
    all_tight = build_all_rollforwards(tb_current, tb_comparative, nominal, materiality=10.0)
    assert all_tight[0].status == "review"


def test_nominal_matrix_other_bucket_materiality_note_respects_configurable_materiality():
    # 11 distinct contra accounts on the same nominal code, each posted a
    # separate transaction (own date) at an equal £100 - MAX_CONTRA_COLUMNS
    # (10) means exactly one of them gets folded into the OTHER bucket
    nominal = pd.DataFrame([
        _nom(f"2025-01-{i:02d}", "8010", "TRADE CREDITORS", credit=100.0, contact="Acme", description=f"inv{i}")
        | {"contra_code": f"70{i:02d}", "contra_name": f"Overhead {i}", "contra_needs_review": False}
        for i in range(1, 12)
    ])
    # loose materiality: the £100 folded into OTHER isn't material enough to
    # call out
    result_loose = build_matrix("8010", "TRADE CREDITORS", nominal, materiality=1000.0)
    assert "folded into" not in result_loose.message
    # tight materiality: the same £100 now clears the bar and gets an
    # explicit advisory note
    result_tight = build_matrix("8010", "TRADE CREDITORS", nominal, materiality=10.0)
    assert "folded into" in result_tight.message


def test_going_concern_respects_configurable_materiality():
    bs = pd.DataFrame([
        {"account_code": "1", "account_name": "Bank", "category": "Bank", "amount": 950.0},
        {"account_code": "2", "account_name": "Creditors", "category": "Current Liability", "amount": -1000.0},
        {"account_code": "3", "account_name": "Share Capital", "category": "Equity", "amount": -50.0},
    ])
    statement = build_bs_statement(bs, net_profit=0.0).statement
    # net current liabilities of £50 - "ok" at the default £500 materiality
    assert going_concern_assess(statement).status == "ok"
    assert going_concern_assess(statement, materiality=10.0).status == "review"


def test_balance_sheet_check_respects_configurable_materiality():
    bs = pd.DataFrame([
        {"account_code": "1", "account_name": "Bank", "category": "Bank", "amount": 1050.0},
        {"account_code": "2", "account_name": "Share Capital", "category": "Equity", "amount": -1000.0},
    ])
    assert build_bs_statement(bs, net_profit=0.0).status == "ok"
    assert build_bs_statement(bs, net_profit=0.0, materiality=10.0).status == "review"


def test_accruals_prepayments_respects_configurable_materiality():
    tb_current = _tb([{"account_code": "620", "account_name": "PREPAYMENTS - INSURANCE", "account_type": "Prepayment", "balance": 1150.0}])
    tb_comparative = _tb([{"account_code": "620", "account_name": "PREPAYMENTS - INSURANCE", "account_type": "Prepayment", "balance": 1000.0}])
    nominal = pd.DataFrame([_nom("2025-06-01", "620", "PREPAYMENTS - INSURANCE", debit=100.0, contact="AXA")])
    # computed c/fwd 1000 + 100 = 1100 vs 1150 per TB: £50 gap
    assert accruals_build_schedule(tb_current, tb_comparative, nominal).status == "ok"
    assert accruals_build_schedule(tb_current, tb_comparative, nominal, materiality=10.0).status == "review"


def test_fixed_asset_category_rollforward_respects_configurable_materiality():
    tb_current = _tb([
        {"account_code": "1", "account_name": "COMPUTER EQUIPMENT COST", "account_type": "Fixed Asset", "balance": 1150.0},
        {"account_code": "2", "account_name": "COMPUTER EQUIPMENT DEPRECIATION", "account_type": "Fixed Asset", "balance": 0.0},
    ])
    tb_comparative = _tb([
        {"account_code": "1", "account_name": "COMPUTER EQUIPMENT COST", "account_type": "Fixed Asset", "balance": 1000.0},
        {"account_code": "2", "account_name": "COMPUTER EQUIPMENT DEPRECIATION", "account_type": "Fixed Asset", "balance": 0.0},
    ])
    nominal = pd.DataFrame([_nom("2025-06-01", "1", "COMPUTER EQUIPMENT COST", debit=100.0)])
    # cost c/fwd computed 1000+100=1100 vs 1150 per TB: £50 gap
    assert category_level_rollforward(tb_current, tb_comparative, nominal).status == "ok"
    assert category_level_rollforward(tb_current, tb_comparative, nominal, materiality=10.0).status == "review"


def test_related_party_transactions_threshold_still_configurable():
    tb_current = _tb([{"account_code": "8100", "account_name": "DIRECTORS LOAN ACCOUNT", "account_type": "Current Liability"}])
    nominal = pd.DataFrame([
        _nom("2025-03-01", "8100", "DIRECTORS LOAN ACCOUNT", credit=3000.0, contact="J Smith"),
        _nom("2025-06-01", "7200", "RENT", debit=100.0, contact="J Smith", description="Rent to director"),
    ])
    assert find_related_party_transactions(tb_current, nominal).status == "ok"
    assert find_related_party_transactions(tb_current, nominal, threshold=50.0).status == "review"


def test_corporation_tax_variance_flag_respects_configurable_materiality():
    result_default = ct_compute(accounting_profit=100000.0, booked_tax_charge=None)
    # give it a booked charge £50 off the computed one
    computed_charge = result_default.tax_charge
    booked = computed_charge + 50.0
    assert ct_compute(accounting_profit=100000.0, booked_tax_charge=booked).status == "ok"
    assert ct_compute(accounting_profit=100000.0, booked_tax_charge=booked, materiality=10.0).status == "review"
