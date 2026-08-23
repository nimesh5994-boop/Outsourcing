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

from app import recon, xero_reports as xr  # noqa: E402
from app.data_sheets import write_data_sheets  # noqa: E402
from app.excel_builder import build_tb_lead_schedule_formulas  # noqa: E402
from app.parsers import FileDataSource  # noqa: E402

SAMPLE_DIR = Path(__file__).resolve().parent.parent / "sample_data"


def _evaluate(path) -> dict:
    xl = formulas.ExcelModel().loads(str(path)).finish()
    return xl.calculate()


def _cell(sol: dict, workbook_filename: str, sheet: str, coord: str):
    key = f"'[{workbook_filename}]{sheet.upper()}'!{coord}"
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
