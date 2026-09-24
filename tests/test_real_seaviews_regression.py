"""Regression test against Seaviews Weymouth Ltd's real, filed source data
- the actual client the ETB/Bank/DLA Activity/Non-current Liabilities/
Prepayments rebuild this session was traced and verified against by hand
(see the session's manual formulas-library verification against the real
downloaded working paper). This automates that same discipline: parse the
real Trial Balance / Account Transactions exports, build the real workbook,
evaluate its live formulas with the `formulas` library (genuine Excel
calculation, not string inspection), and cross-check every figure against
an independently-computed ground truth from the same real source data -
so a future change that breaks one of these sections fails a test instead
of only showing up on the next manual production check.

The real files live only in the user's OneDrive (never committed to the
repo - this is a real client's real financial data), so this whole module
is skipped anywhere they aren't present, e.g. CI or another machine."""
from pathlib import Path

import openpyxl
import pytest

import formulas

from app import (
    accruals_prepayments, compliance_checks, control_accounts, corporation_tax,
    excel_builder, fixed_assets, nominal_matrix, parsers, recon, xero_reports,
)
from app.excel_builder import build_workbook

REAL_DIR = Path(
    r"C:\Users\Nimesh\OneDrive - Affinity Associates Ltd\Documents\CAPTO APP CONTENT\WP auto setup\WP_sample"
)

pytestmark = pytest.mark.skipif(
    not REAL_DIR.exists(),
    reason="Real Seaviews Weymouth Ltd sample files only available on this dev machine - never committed to the repo.",
)

CLIENT_NAME = "Seaviews Weymouth Ltd"
CURRENT_LABEL = "01 Apr 2025 to 31 Mar 2026"
COMPARATIVE_LABEL = "01 Apr 2024 to 31 Mar 2025"


@pytest.fixture(scope="module")
def real_seaviews_data():
    tb_current, _ = xero_reports.parse_trial_balance(
        parsers.FileDataSource(REAL_DIR / "Seaviews_Weymouth_Ltd_-_Trial_Balance 2026.xlsx")
    )
    tb_comparative, _ = xero_reports.parse_trial_balance(
        parsers.FileDataSource(REAL_DIR / "Seaviews_Weymouth_Ltd_-_Trial_Balance 25.xlsx")
    )
    nominal_current = xero_reports.parse_account_transactions(
        parsers.FileDataSource(REAL_DIR / "Seaviews_Weymouth_Ltd_-_Account_Transactions 2026.xlsx")
    )
    nominal_comparative = xero_reports.parse_account_transactions(
        parsers.FileDataSource(REAL_DIR / "Seaviews_Weymouth_Ltd_-_Account_Transactions 2025.xlsx")
    )
    pl_current, bs_current = xero_reports.derive_pl_bs_from_tb(tb_current)
    return {
        "tb_current": tb_current, "tb_comparative": tb_comparative,
        "nominal_current": nominal_current, "nominal_comparative": nominal_comparative,
        "pl_current": pl_current, "bs_current": bs_current,
    }


@pytest.fixture(scope="module")
def real_seaviews_workbook_path(tmp_path_factory, real_seaviews_data):
    data = real_seaviews_data
    results = recon.run_all_recons(data)
    # accruals_prepayments.build_schedule is computed inline as part of the
    # real generate() pipeline's own results list (app/main.py ~line 1908),
    # not inside recon.run_all_recons - omitting it here would silently skip
    # the whole Prepayments & Accrued Income schedule in the built workbook.
    results = results + [accruals_prepayments.build_schedule(
        data.get("tb_current"), data.get("tb_comparative"), data.get("nominal_current"), 500.0,
    )]
    ca_results = control_accounts.build_all_rollforwards(
        data["tb_current"], data["tb_comparative"], data["nominal_current"], None, None,
    )
    mx_results = nominal_matrix.build_all_matrices(data["tb_current"], data["nominal_current"])
    ct_computation = corporation_tax.compute(accounting_profit=float(data["pl_current"]["amount"].sum()))
    fixed_asset_result = fixed_assets.category_level_rollforward(
        data["tb_current"], data["tb_comparative"], data["nominal_current"]
    )

    wb = build_workbook(
        CLIENT_NAME, CURRENT_LABEL, COMPARATIVE_LABEL, data, results,
        control_account_results=ca_results, matrix_results=mx_results,
        ct_computation=ct_computation, fixed_asset_result=fixed_asset_result,
    )
    out_dir = tmp_path_factory.mktemp("real_seaviews")
    out = out_dir / "Real_Seaviews_regression.xlsx"
    wb.save(out)
    return out


def _evaluate(workbook_path: Path):
    """Loads the workbook into the `formulas` library and returns a lookup
    function - the same genuine-calculation approach (not string inspection)
    used throughout this session's manual production verification."""
    xl_model = formulas.ExcelModel().loads(str(workbook_path)).finish()
    sol = xl_model.calculate()
    fname = workbook_path.name

    def get(sheet: str, cell: str):
        key = f"'[{fname}]{sheet.upper()}'!{cell}"
        value = sol[key].value
        try:
            return value[0][0]
        except Exception:
            return value

    return get


def _find_row_by_label(ws, label_column: str, label_text: str) -> int:
    for row in ws.iter_rows():
        cell = row[openpyxl.utils.column_index_from_string(label_column) - 1]
        if cell.value == label_text:
            return cell.row
    raise AssertionError(f"No row found with {label_column}={label_text!r} on sheet {ws.title!r}")


def _find_row_by_account_code(ws, code_column: str, code: str) -> int:
    for row in ws.iter_rows():
        cell = row[openpyxl.utils.column_index_from_string(code_column) - 1]
        if str(cell.value) == code:
            return cell.row
    raise AssertionError(f"No row found with {code_column}={code!r} on sheet {ws.title!r}")


def test_real_trial_balance_self_balances(real_seaviews_data):
    assert abs(real_seaviews_data["tb_current"]["balance"].sum()) < 0.01
    assert abs(real_seaviews_data["tb_comparative"]["balance"].sum()) < 0.01


def test_real_workbook_builds_with_every_section_this_session_added(real_seaviews_workbook_path):
    wb = openpyxl.load_workbook(real_seaviews_workbook_path)
    assert any(s.startswith("2 ETB") or s == "2 ETB" for s in wb.sheetnames)
    assert any("Bank" in s for s in wb.sheetnames)
    assert any("DLA Activity" in s for s in wb.sheetnames)
    assert any(s.endswith("Prepayments") for s in wb.sheetnames)


def test_real_etb_xero_ptb_totals_match_the_real_trial_balance(real_seaviews_data, real_seaviews_workbook_path):
    """The ETB's "Xero PTB" columns must reproduce the real, raw, unadjusted
    TB exactly - this is the actual mechanism traced from the real finished
    WP file this session (Draft = Xero PTB, before any journal adjustment)."""
    wb = openpyxl.load_workbook(real_seaviews_workbook_path, data_only=False)
    etb_ws = next(ws for ws in wb.worksheets if ws.title.startswith("2 ETB") or ws.title == "2 ETB")
    get = _evaluate(real_seaviews_workbook_path)

    tb_current = real_seaviews_data["tb_current"]
    expected_dr_total = float(tb_current["debit"].sum())
    expected_cr_total = float(tb_current["credit"].sum())

    header_row = _find_row_by_label(etb_ws, "A", "Account Code")
    dr_col = cr_col = None
    for cell in etb_ws[header_row]:
        if cell.value == "Xero PTB Dr":
            dr_col = cell.column_letter
        elif cell.value == "Xero PTB Cr":
            cr_col = cell.column_letter
    assert dr_col and cr_col, "ETB header row must expose separate Xero PTB Dr/Cr columns"

    actual_dr_total = actual_cr_total = 0.0
    for row in etb_ws.iter_rows(min_row=header_row + 1):
        code_cell = row[0]
        if code_cell.value is None:
            continue
        dr_val = get(etb_ws.title, f"{dr_col}{code_cell.row}")
        cr_val = get(etb_ws.title, f"{cr_col}{code_cell.row}")
        actual_dr_total += float(dr_val or 0)
        actual_cr_total += float(cr_val or 0)

    assert actual_dr_total == pytest.approx(expected_dr_total, abs=0.01)
    assert actual_cr_total == pytest.approx(expected_cr_total, abs=0.01)


def test_real_dla_activity_ties_to_the_trial_balance_with_zero_difference(real_seaviews_data, real_seaviews_workbook_path):
    """The DLA is confirmed present for this real client (compliance_checks
    finds it via _DLA_PATTERN) - its running balance must close to exactly
    the real TB's own balance for that account, the same zero-diff check
    this session verified by hand against production."""
    dla_accounts = compliance_checks._find_accounts(real_seaviews_data["tb_current"], compliance_checks._DLA_PATTERN)
    assert not dla_accounts.empty, "Seaviews' real TB is expected to carry a Directors' Loan Account"

    wb = openpyxl.load_workbook(real_seaviews_workbook_path, data_only=False)
    dla_ws = next(ws for ws in wb.worksheets if "DLA Activity" in ws.title)
    get = _evaluate(real_seaviews_workbook_path)

    diff_row = _find_row_by_label(dla_ws, "D", "Diff (should be nil)")
    diff = get(dla_ws.title, f"G{diff_row}")
    assert diff == pytest.approx(0.0, abs=0.01)

    code = str(dla_accounts.iloc[0]["account_code"])
    expected_closing = float(
        real_seaviews_data["tb_current"].loc[
            real_seaviews_data["tb_current"]["account_code"].astype(str) == code, "balance"
        ].sum()
    )
    closing_row = _find_row_by_label(dla_ws, "D", "Closing balance per this ledger")
    actual_closing = get(dla_ws.title, f"G{closing_row}")
    assert actual_closing == pytest.approx(expected_closing, abs=0.01)


def test_real_prepayments_accruals_schedule_matches_the_real_trial_balance(real_seaviews_data, real_seaviews_workbook_path):
    """Every prepayment/accrual line's Draft column must equal that account's
    own raw TB balance, and (with no journal posted yet on a fresh generate)
    Total must equal Draft - the same £834.00 accrued figure this session
    matched by hand to the live recon summary."""
    accounts = accruals_prepayments._find_accounts(real_seaviews_data["tb_current"])
    assert accounts, "Seaviews' real TB is expected to carry at least one prepayment/accrual account"

    wb = openpyxl.load_workbook(real_seaviews_workbook_path, data_only=False)
    prepay_ws = next(ws for ws in wb.worksheets if ws.title.endswith("Prepayments") and "Control" not in "".join(
        str(c.value) for c in ws[3] if c.value
    ))
    get = _evaluate(real_seaviews_workbook_path)

    tb_current = real_seaviews_data["tb_current"]
    for code, name, _account_type in accounts:
        row = _find_row_by_account_code(prepay_ws, "B", code)
        draft = get(prepay_ws.title, f"D{row}")
        adjustment = get(prepay_ws.title, f"E{row}")
        total = get(prepay_ws.title, f"F{row}")

        expected_draft = float(
            tb_current.loc[tb_current["account_code"].astype(str) == code, "debit"].sum()
            - tb_current.loc[tb_current["account_code"].astype(str) == code, "credit"].sum()
        )
        assert draft == pytest.approx(expected_draft, abs=0.01), f"account {code} ({name}) draft mismatch"
        assert adjustment == pytest.approx(0.0, abs=0.01)  # blank Journals sheet on a fresh generate
        assert total == pytest.approx(expected_draft, abs=0.01)


def test_real_bank_lead_schedule_lists_every_real_bank_account(real_seaviews_data, real_seaviews_workbook_path):
    tb_current = real_seaviews_data["tb_current"]
    bank_accounts = tb_current[tb_current["account_type"].astype(str).str.lower() == "bank"]
    assert not bank_accounts.empty, "Seaviews' real TB is expected to carry at least one Bank account"

    wb = openpyxl.load_workbook(real_seaviews_workbook_path, data_only=False)
    bank_ws = next(ws for ws in wb.worksheets if ws.title.split(" ", 1)[-1] == "Bank")
    get = _evaluate(real_seaviews_workbook_path)

    header_row = _find_row_by_label(bank_ws, "A", "Account Code")
    for _, acc in bank_accounts.iterrows():
        code = str(acc["account_code"])
        row = _find_row_by_account_code(bank_ws, "A", code)
        assert row > header_row
        draft = get(bank_ws.title, f"C{row}")
        expected_draft = float(acc["debit"] - acc["credit"])
        assert draft == pytest.approx(expected_draft, abs=0.01)
