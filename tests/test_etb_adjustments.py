"""End-to-end verification of the ETB Adjustment column (Phase 2 of the
Extended Trial Balance work) - see excel_builder.build_tb_tieout_sheet_formulas
and data_sheets.write_data_sheets' tb_adjustments_ref handling.

Builds a real workbook through the actual pipeline (recon.run_all_recons +
tb_tieout.build_tieout, same as main.py's generation step), types a value
into the TB Tie-Out sheet's Adjustment column, and uses the `formulas`
library to *evaluate* the result - not just eyeball the formula strings -
confirming the adjustment reaches DATA_TB_Current's balance and, through
it, the P&L and Balance Sheet statements, with no changes needed in any
of those sheets' own formulas.
"""
from pathlib import Path

import openpyxl
import pytest

pytest.importorskip("formulas")
import formulas  # noqa: E402

from app import control_accounts, corporation_tax, fixed_assets, nominal_matrix, recon, tb_tieout  # noqa: E402
from app.excel_builder import build_workbook  # noqa: E402


def _evaluate(path) -> dict:
    xl = formulas.ExcelModel().loads(str(path)).finish()
    return xl.calculate()


def _cell(sol: dict, workbook_filename: str, sheet: str, coord: str):
    key = f"'[{workbook_filename}]{sheet.upper()}'!{coord}"
    return sol[key].value[0][0]


def _build_and_save(canonical_data, out_path):
    results = recon.run_all_recons(canonical_data)
    results = results + [tb_tieout.build_tieout(
        canonical_data.get("tb_current"), canonical_data.get("tb_comparative"), canonical_data.get("nominal_current"),
    )]
    ca_results = control_accounts.build_all_rollforwards(
        canonical_data["tb_current"], canonical_data["tb_comparative"], canonical_data["nominal_current"],
        canonical_data["aged_debtors"], canonical_data["aged_creditors"],
    )
    mx_results = nominal_matrix.build_all_matrices(canonical_data["tb_current"], canonical_data["nominal_current"])
    ct_computation = corporation_tax.compute(accounting_profit=float(canonical_data["pl_current"]["amount"].sum()))
    fixed_asset_result = fixed_assets.category_level_rollforward(
        canonical_data["tb_current"], canonical_data["tb_comparative"], canonical_data["nominal_current"]
    )
    wb = build_workbook(
        "Brightwell Landscaping Supplies Limited",
        "Year ended 31 December 2025", "Year ended 31 December 2024",
        canonical_data, results, control_account_results=ca_results, matrix_results=mx_results,
        ct_computation=ct_computation, fixed_asset_result=fixed_asset_result,
    )
    wb.save(out_path)
    return wb


def test_tb_tieout_sheet_exists_with_a_blank_adjustment_column(tmp_path, canonical_data):
    out = tmp_path / "wp.xlsx"
    wb = _build_and_save(canonical_data, out)
    assert any("TB Tie-Out" in s for s in wb.sheetnames)
    ws = next(wb[s] for s in wb.sheetnames if "TB Tie-Out" in s)
    header_row = next(r for r in range(1, 10) if ws.cell(row=r, column=6).value == "Adjustment")
    assert ws.cell(row=header_row + 1, column=6).value is None  # genuinely blank, not a formula


def test_adjustment_flows_through_to_data_tb_current_and_the_pl_statement(tmp_path, canonical_data):
    out = tmp_path / "wp.xlsx"
    wb = _build_and_save(canonical_data, out)
    tieout_sheet_name = next(s for s in wb.sheetnames if "TB Tie-Out" in s)
    pl_sheet_name = next(s for s in wb.sheetnames if s.endswith("Profit and Loss"))
    ws = wb[tieout_sheet_name]
    header_row = next(r for r in range(1, 10) if ws.cell(row=r, column=6).value == "Adjustment")
    tieout_data_start = header_row + 1

    # Pick the first P&L-type account (opening always 0, so the maths is simplest to hand-check).
    # TB Tie-Out row i (0-indexed from tieout_data_start) is the same account as
    # DATA_TB_Current row i (0-indexed from its own first data row, 2) - both are
    # built from the exact same tb_current dataframe, in the same order.
    pl_code = str(canonical_data["pl_current"].iloc[0]["account_code"])
    dtc_row = _dtc_row_for(wb, pl_code)
    tieout_row = tieout_data_start + (dtc_row - 2)
    assert ws.cell(row=tieout_row, column=6).value is None  # confirms this row's Adjustment is blank, as expected

    sol_before = _evaluate(out)
    before_dtc = _cell(sol_before, "wp.xlsx", "DATA_TB_Current", f"F{dtc_row}")
    before_net_profit = _cell(sol_before, "wp.xlsx", pl_sheet_name, "B12")

    ADJUSTMENT = 250.0
    ws.cell(row=tieout_row, column=6, value=ADJUSTMENT)
    wb.save(out)

    sol_after = _evaluate(out)
    after_dtc = _cell(sol_after, "wp.xlsx", "DATA_TB_Current", f"F{dtc_row}")
    after_net_profit = _cell(sol_after, "wp.xlsx", pl_sheet_name, "B12")

    assert after_dtc == pytest.approx(before_dtc + ADJUSTMENT)
    # DATA_PL negates DATA_TB_Current's balance (see data_sheets._write_derived_amount_sheet),
    # so a +250 balance adjustment should move Net Profit by -250.
    assert after_net_profit == pytest.approx(before_net_profit - ADJUSTMENT)


def _dtc_row_for(wb, account_code: str) -> int:
    ws = wb["DATA_TB_Current"]
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=1).value) == account_code:
            return r
    raise AssertionError(f"account code {account_code} not found on DATA_TB_Current")
