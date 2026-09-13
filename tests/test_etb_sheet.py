"""End-to-end verification of the ETB sheet (see excel_builder.build_etb_sheet_formulas)
- the consolidated Extended Trial Balance built *in addition to* the TB
Tie-Out sheet, per the user's explicit "keep what is there but build ETB
sheet too" instruction. Covers what makes it different from TB Tie-Out:
a Comparative Closing reference column, and its own independent
Adjustment column that sums with TB Tie-Out's rather than needing to
agree with it.
"""
from pathlib import Path

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


def _dtc_row_for(wb, account_code: str) -> int:
    ws = wb["DATA_TB_Current"]
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=1).value) == account_code:
            return r
    raise AssertionError(f"account code {account_code} not found on DATA_TB_Current")


def test_etb_sheet_exists_alongside_tb_tieout_with_its_own_blank_adjustment_column(tmp_path, canonical_data):
    out = tmp_path / "wp.xlsx"
    wb = _build_and_save(canonical_data, out)
    assert any("TB Tie-Out" in s for s in wb.sheetnames)  # kept, per the user's instruction
    assert any(s.endswith("ETB") for s in wb.sheetnames)  # added

    ws = next(wb[s] for s in wb.sheetnames if s.endswith("ETB"))
    header_row = next(r for r in range(1, 12) if ws.cell(row=r, column=1).value == "Account Code")
    headers = [ws.cell(row=header_row, column=c).value for c in range(1, 13)]
    assert headers == [
        "Account Code", "Account Name", "Account Type",
        "Comparative Closing (per comparative TB)", "Opening (per comparative TB)",
        "Movement (current year)", "Adjustment", "Derived Closing",
        "Reported Closing (per current TB)", "Diff", "Flag", "Notes",
    ]
    assert ws.cell(row=header_row + 1, column=7).value is None  # Adjustment (G) genuinely blank
    assert ws.cell(row=header_row + 1, column=12).value is None  # Notes (L) genuinely blank


def test_comparative_closing_is_not_zeroed_for_a_pl_account_but_opening_is(tmp_path, canonical_data):
    # Opening must be 0 for a P&L account (same reasoning as tb_tieout.py -
    # an income statement account doesn't carry a balance forward), but
    # Comparative Closing is a plain reference figure and should show last
    # year's real P&L total for that code, not be zeroed the same way.
    out = tmp_path / "wp.xlsx"
    wb = _build_and_save(canonical_data, out)
    ws = next(wb[s] for s in wb.sheetnames if s.endswith("ETB"))
    header_row = next(r for r in range(1, 12) if ws.cell(row=r, column=1).value == "Account Code")
    data_start = header_row + 1

    pl_code = str(canonical_data["pl_current"].iloc[0]["account_code"])
    dtc_row = _dtc_row_for(wb, pl_code)
    etb_row = data_start + (dtc_row - 2)

    assert ws.cell(row=etb_row, column=5).value == "=0"  # Opening
    comparative_closing_formula = ws.cell(row=etb_row, column=4).value
    assert comparative_closing_formula != "=0"  # Comparative Closing - a real lookup, not zeroed

    sol = _evaluate(out)
    etb_sheet_name = next(s for s in wb.sheetnames if s.endswith("ETB"))
    comparative_closing_value = _cell(sol, "wp.xlsx", etb_sheet_name, f"D{etb_row}")
    expected = float(canonical_data["tb_comparative"].groupby("account_code")["balance"].sum().get(pl_code, 0.0))
    assert comparative_closing_value == pytest.approx(expected)


def test_etb_adjustment_is_independent_of_tb_tieout_and_both_sum_into_data_tb_current(tmp_path, canonical_data):
    out = tmp_path / "wp.xlsx"
    wb = _build_and_save(canonical_data, out)

    tieout_sheet_name = next(s for s in wb.sheetnames if "TB Tie-Out" in s)
    etb_sheet_name = next(s for s in wb.sheetnames if s.endswith("ETB"))
    tieout_ws, etb_ws = wb[tieout_sheet_name], wb[etb_sheet_name]

    tieout_header_row = next(r for r in range(1, 10) if tieout_ws.cell(row=r, column=6).value == "Adjustment")
    tieout_data_start = tieout_header_row + 1
    etb_header_row = next(r for r in range(1, 12) if etb_ws.cell(row=r, column=1).value == "Account Code")
    etb_data_start = etb_header_row + 1

    pl_code = str(canonical_data["pl_current"].iloc[0]["account_code"])
    dtc_row = _dtc_row_for(wb, pl_code)
    tieout_row = tieout_data_start + (dtc_row - 2)
    etb_row = etb_data_start + (dtc_row - 2)

    sol_before = _evaluate(out)
    before_dtc = _cell(sol_before, "wp.xlsx", "DATA_TB_Current", f"F{dtc_row}")

    TIEOUT_ADJUSTMENT, ETB_ADJUSTMENT = 250.0, 60.0
    tieout_ws.cell(row=tieout_row, column=6, value=TIEOUT_ADJUSTMENT)
    etb_ws.cell(row=etb_row, column=7, value=ETB_ADJUSTMENT)
    wb.save(out)

    sol_after = _evaluate(out)
    after_dtc = _cell(sol_after, "wp.xlsx", "DATA_TB_Current", f"F{dtc_row}")

    # Both adjustments land on the same account, from two different sheets,
    # and simply add - neither is required to agree with (or know about) the other.
    assert after_dtc == pytest.approx(before_dtc + TIEOUT_ADJUSTMENT + ETB_ADJUSTMENT)
