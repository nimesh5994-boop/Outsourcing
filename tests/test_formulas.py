"""Verifies formula-linked schedules actually evaluate correctly - not just
that the string looks like a formula, but that a real formula engine
computes the same numbers the equivalent Python (pandas) computation does.

Uses the `formulas` library (pure-Python Excel formula evaluator) since
LibreOffice can't be relied on in this sandbox to recalculate and verify.
These tests are skipped if `formulas` isn't installed - it's a heavy,
dev-only dependency (requirements-dev.txt), not part of the app's runtime
requirements.
"""
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook

pytest.importorskip("formulas")
import formulas  # noqa: E402

from app import control_accounts as ca  # noqa: E402
from app import recon, xero_reports as xr  # noqa: E402
from app.data_sheets import write_data_sheets  # noqa: E402
from app.excel_builder import build_control_account_sheet_formulas, build_tb_lead_schedule_formulas  # noqa: E402
from app.parsers import FileDataSource  # noqa: E402

SAMPLE_DIR = Path(__file__).resolve().parent.parent / "sample_data"


def _evaluate(path) -> dict:
    xl = formulas.ExcelModel().loads(str(path)).finish()
    return xl.calculate()


def _cell(sol: dict, workbook_filename: str, sheet: str, coord: str):
    key = f"'[{workbook_filename}]{sheet.upper()}'!{coord}"
    return sol[key].value[0][0]


def _cell_or_none(sol: dict, workbook_filename: str, sheet: str, coord: str):
    key = f"'[{workbook_filename}]{sheet.upper()}'!{coord}"
    if key not in sol:
        return None
    return sol[key].value[0][0]


@pytest.fixture
def canonical_tb():
    tb_current, tb_comparative = xr.parse_trial_balance(
        FileDataSource(SAMPLE_DIR / "trial_balance_current_xero.xlsx")
    )
    return tb_current, tb_comparative


def test_tb_lead_schedule_formulas_match_python_ground_truth(tmp_path, canonical_tb):
    tb_current, tb_comparative = canonical_tb
    variance = recon.variance_analysis(tb_current, tb_comparative)

    wb = Workbook()
    refs = write_data_sheets(wb, {"tb_current": tb_current, "tb_comparative": tb_comparative})
    build_tb_lead_schedule_formulas(wb, "Brightwell Landscaping Supplies Limited", "Year ended 31 December 2025", "1", variance.detail, refs)
    wb.remove(wb["Sheet"])

    out = tmp_path / "tb_formula.xlsx"
    wb.save(out)

    sol = _evaluate(out)
    sheet_name = "1 TB Lead Schedule"

    errors = [k for k, v in sol.items() if f"[{out.name}]{sheet_name.upper()}" in k and "VALUE!" in str(v.value)]
    assert not errors, f"formula errors: {errors}"

    for i, row in enumerate(variance.detail.itertuples()):
        r = 6 + i
        assert _cell(sol, out.name, sheet_name, f"C{r}") == pytest.approx(row.current_year, abs=0.01)
        assert _cell(sol, out.name, sheet_name, f"D{r}") == pytest.approx(row.comparative_year, abs=0.01)
        assert _cell(sol, out.name, sheet_name, f"E{r}") == pytest.approx(row.variance_amount, abs=0.01)
        assert bool(_cell(sol, out.name, sheet_name, f"G{r}")) == bool(row.flag)


def test_control_account_formulas_match_python_ground_truth(tmp_path, canonical_data):
    results = ca.build_all_rollforwards(
        canonical_data["tb_current"], canonical_data["tb_comparative"], canonical_data["nominal_current"],
        canonical_data["aged_debtors"], canonical_data["aged_creditors"],
    )
    trade_creditors = next(r for r in results if r.account_code == "8010")

    wb = Workbook()
    refs = write_data_sheets(wb, canonical_data)
    build_control_account_sheet_formulas(wb, "Brightwell Landscaping Supplies Limited", "Year ended 31 December 2025", "1", trade_creditors, refs)
    wb.remove(wb["Sheet"])

    out = tmp_path / "ca_formula.xlsx"
    wb.save(out)

    sol = _evaluate(out)
    sheet_name = f"1 {trade_creditors.account_name}"[:31]

    errors = [k for k, v in sol.items() if f"[{out.name}]{sheet_name.upper()}" in k and ("VALUE!" in str(v.value) or "REF!" in str(v.value))]
    assert not errors, f"formula errors: {errors}"

    # BALANCE C/FWD (per TB) row: Debit/Credit split of the current TB balance
    schedule_rows = {row["Item"]: i for i, row in trade_creditors.schedule.iterrows()}
    c_fwd_row_idx = schedule_rows["BALANCE C/FWD (per TB)"]
    expected_debit = trade_creditors.schedule.iloc[c_fwd_row_idx]["Debit £"]
    expected_credit = trade_creditors.schedule.iloc[c_fwd_row_idx]["Credit £"]
    excel_row = 8 + c_fwd_row_idx  # row 5 = status, table header at row 7, first data row 8
    assert _cell(sol, out.name, sheet_name, f"C{excel_row}") == pytest.approx(expected_debit, abs=0.01)
    assert _cell(sol, out.name, sheet_name, f"D{excel_row}") == pytest.approx(expected_credit, abs=0.01)

    # the breakdown's "unexplained difference" should be ~0 on this data (confirmed against real client data separately)
    for r in range(6, 6 + len(trade_creditors.schedule) + len(trade_creditors.breakdown) + 6):
        label = _cell_or_none(sol, out.name, sheet_name, f"A{r}")
        if label and "UNEXPLAINED" in str(label):
            # this dataset's aged creditors listing doesn't fully explain the
            # TB balance (a known, deliberate gap - see
            # test_control_account_breakdown_flags_unexplained_difference in
            # test_pipeline.py), so just confirm the formula reproduces the
            # same gap the Python computation found, not that it's zero
            unexplained_row = trade_creditors.breakdown[
                trade_creditors.breakdown["Party"].astype(str).str.startswith("UNEXPLAINED")
            ].iloc[0]
            assert _cell(sol, out.name, sheet_name, f"B{r}") == pytest.approx(unexplained_row["Amount £"], abs=0.5)
            break
    else:
        pytest.fail("did not find the UNEXPLAINED DIFFERENCE row")
