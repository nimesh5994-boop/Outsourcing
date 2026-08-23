"""End-to-end test: parse the sample data (mirroring real Xero export
structures), run every reconciliation, and build the final workbook."""
from pathlib import Path

import openpyxl
import pytest

from app import control_accounts, mapping, nominal_matrix, parsers, recon, xero_reports
from app.excel_builder import build_workbook

SAMPLE_DIR = Path(__file__).resolve().parent.parent / "sample_data"


@pytest.fixture
def canonical_data():
    tb_current, tb_comparative = xero_reports.parse_trial_balance(
        parsers.FileDataSource(SAMPLE_DIR / "trial_balance_current_xero.xlsx")
    )
    nominal_current = xero_reports.parse_account_transactions(
        parsers.FileDataSource(SAMPLE_DIR / "account_transactions_current_xero.xlsx")
    )
    aged_debtors = xero_reports.parse_aged_report(
        parsers.FileDataSource(SAMPLE_DIR / "aged_receivables_current_xero.xlsx"), "customer"
    )
    aged_creditors = xero_reports.parse_aged_report(
        parsers.FileDataSource(SAMPLE_DIR / "aged_payables_current_xero.xlsx"), "supplier"
    )
    pl_current, bs_current = xero_reports.derive_pl_bs_from_tb(tb_current)

    vat_source = parsers.FileDataSource(SAMPLE_DIR / "vat_return_current.csv")
    vat_mapping = mapping.suggest_mapping("vat_return", vat_source.raw_columns())
    vat_return = parsers.apply_mapping(vat_source, "vat_return", vat_mapping)

    bank_source = parsers.FileDataSource(SAMPLE_DIR / "bank_statement_current.csv")
    bank_mapping = mapping.suggest_mapping("bank_statement", bank_source.raw_columns())
    bank_statement = parsers.apply_mapping(bank_source, "bank_statement", bank_mapping)

    return {
        "tb_current": tb_current, "tb_comparative": tb_comparative,
        "nominal_current": nominal_current,
        "aged_debtors": aged_debtors, "aged_creditors": aged_creditors,
        "pl_current": pl_current, "bs_current": bs_current,
        "vat_return": vat_return, "bank_statement": bank_statement,
    }


def test_trial_balance_self_balances(canonical_data):
    assert abs(canonical_data["tb_current"]["balance"].sum()) < 0.01
    assert abs(canonical_data["tb_comparative"]["balance"].sum()) < 0.01


def test_aged_report_sums_from_detail_rows_not_formula_subtotals(canonical_data):
    # the sample files use un-evaluated '=E9'-style formulas in the subtotal
    # rows (matching Xero's real export quirk) - a naive parser would read
    # those as zero
    moss = canonical_data["aged_debtors"].set_index("customer").loc["MOSS & TURF SUPPLIES"]
    assert moss["total"] == pytest.approx(2370.00)


def test_vat_column_mapping_via_alias_suggestion(canonical_data):
    v = canonical_data["vat_return"].iloc[0]
    assert v["box5"] == pytest.approx(8380.00)
    assert v["box6"] == pytest.approx(78400.00)


def test_recon_engine_runs_all_checks(canonical_data):
    results = recon.run_all_recons(canonical_data)
    names = {r.name for r in results}
    assert "TB self-balance check" in names
    assert all(r.status in ("ok", "review", "error", "n/a") for r in results)
    tb_check = next(r for r in results if r.name == "TB self-balance check")
    assert tb_check.status == "ok"


def test_control_account_rollforwards_compute_correctly(canonical_data):
    # the hand-built sample nominal activity is an illustrative subset, not a
    # complete GL, so accounts won't all tie out perfectly - what matters is
    # that b/fwd + movements - c/fwd is computed correctly either way
    results = control_accounts.build_all_rollforwards(
        canonical_data["tb_current"], canonical_data["tb_comparative"], canonical_data["nominal_current"],
        canonical_data["aged_debtors"], canonical_data["aged_creditors"],
    )
    assert len(results) > 0
    assert all(r.status in ("ok", "review") for r in results)
    for r in results:
        b_fwd = r.schedule.iloc[0]["Debit £"] - r.schedule.iloc[0]["Credit £"]
        move_debit, move_credit = r.schedule.iloc[1]["Debit £"], r.schedule.iloc[1]["Credit £"]
        c_fwd = r.schedule.iloc[2]["Debit £"] - r.schedule.iloc[2]["Credit £"]
        computed = b_fwd + move_debit - move_credit
        rollforward_ties = abs(computed - c_fwd) <= control_accounts.MATERIALITY_AMOUNT
        if not rollforward_ties:
            assert r.status == "review"


def test_control_account_breakdown_flags_unexplained_difference(canonical_data):
    # accounts receivable's balance (96,420) is far bigger than the tiny
    # sample aged debtors listing (15,170) - the breakdown should surface
    # that gap as a single explicit number rather than silently ignore it
    results = control_accounts.build_all_rollforwards(
        canonical_data["tb_current"], canonical_data["tb_comparative"], canonical_data["nominal_current"],
        canonical_data["aged_debtors"], canonical_data["aged_creditors"],
    )
    receivable = next(r for r in results if r.account_code == "610A")
    assert not receivable.breakdown.empty
    unexplained_row = receivable.breakdown[receivable.breakdown["Party"].str.startswith("UNEXPLAINED")].iloc[0]
    assert unexplained_row["Amount £"] == pytest.approx(15170.00 - 96420.00)
    assert receivable.status == "review"


def test_nominal_matrix_flags_multi_code_splits(canonical_data):
    results = nominal_matrix.build_all_matrices(
        canonical_data["tb_current"], canonical_data["nominal_current"],
        account_codes=["8010"],
    )
    trade_creditors = results[0]
    assert trade_creditors.status == "review"
    assert "multi-code split" in trade_creditors.message


def test_full_workbook_builds_and_saves(tmp_path, canonical_data):
    results = recon.run_all_recons(canonical_data)
    ca_results = control_accounts.build_all_rollforwards(
        canonical_data["tb_current"], canonical_data["tb_comparative"], canonical_data["nominal_current"],
        canonical_data["aged_debtors"], canonical_data["aged_creditors"],
    )
    mx_results = nominal_matrix.build_all_matrices(canonical_data["tb_current"], canonical_data["nominal_current"])

    wb = build_workbook(
        "Brightwell Landscaping Supplies Limited",
        "Year ended 31 December 2025", "Year ended 31 December 2024",
        canonical_data, results, control_account_results=ca_results, matrix_results=mx_results,
    )
    out = tmp_path / "working_paper.xlsx"
    wb.save(out)

    assert out.exists()
    reopened = openpyxl.load_workbook(out)
    assert "Index" in reopened.sheetnames
    assert any("TB Lead Schedule" in s for s in reopened.sheetnames)
    assert len(reopened.sheetnames) > 10
