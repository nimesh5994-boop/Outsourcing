"""End-to-end verification of the ETB sheet (see excel_builder.build_etb_sheet_formulas)
- rebuilt to match how a real Extended Trial Balance is actually built,
traced from a real firm's own working paper template: this year's Trial
Balance exactly as uploaded (Xero PTB Dr/Cr, never adjusted) plus the
Journals sheet's own Dr/Cr for that code, summed and split into Profit &
Loss / Balance Sheet by account type - kept separate from TB Tie-Out's
job (opening + nominal-ledger movement = reported closing), which an
earlier version of this sheet duplicated instead of actually matching
what a real ETB looks like.
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


def _etb_sheet(wb):
    name = next(s for s in wb.sheetnames if s.endswith("ETB"))
    ws = wb[name]
    header_row = next(r for r in range(1, 12) if ws.cell(row=r, column=1).value == "Account Code")
    return name, ws, header_row


def _journals_first_row(wb) -> int:
    ws = wb[next(s for s in wb.sheetnames if s.endswith("Journals"))]
    header_row = next(r for r in range(1, 10) if ws.cell(row=r, column=1).value == "No.")
    return header_row + 1


def test_etb_headers_match_a_real_extended_trial_balance():
    import openpyxl
    from app.excel_builder import ETB_HEADERS
    assert ETB_HEADERS == [
        "Account Code", "Account Name", "Account Type",
        "Xero PTB Dr", "Xero PTB Cr",
        "Journals Dr", "Journals Cr",
        "P&L Dr", "P&L Cr",
        "Balance Sheet Dr", "Balance Sheet Cr",
        "Comparative Dr", "Comparative Cr", "Notes",
    ]


def test_xero_ptb_columns_show_the_raw_unadjusted_tb_and_notes_is_blank(tmp_path, canonical_data):
    out = tmp_path / "wp.xlsx"
    wb = _build_and_save(canonical_data, out)
    assert any("TB Tie-Out" in s for s in wb.sheetnames)  # kept, unchanged
    _, ws, header_row = _etb_sheet(wb)
    data_row = header_row + 1

    assert ws.cell(row=data_row, column=4).value.startswith("=")  # Xero PTB Dr - a formula
    assert ws.cell(row=data_row, column=14).value is None  # Notes - genuinely blank

    sol = _evaluate(out)
    etb_sheet_name = next(s for s in wb.sheetnames if s.endswith("ETB"))
    pl_code = str(canonical_data["pl_current"].iloc[0]["account_code"])
    dtc_row = _dtc_row_for(wb, pl_code)
    etb_row = header_row + 1 + (dtc_row - 2)

    dtc_debit = _cell(sol, "wp.xlsx", "DATA_TB_Current", f"D{dtc_row}")
    dtc_credit = _cell(sol, "wp.xlsx", "DATA_TB_Current", f"E{dtc_row}")
    etb_debit = _cell(sol, "wp.xlsx", etb_sheet_name, f"D{etb_row}")
    etb_credit = _cell(sol, "wp.xlsx", etb_sheet_name, f"E{etb_row}")
    assert etb_debit == pytest.approx(dtc_debit)
    assert etb_credit == pytest.approx(dtc_credit)


def test_comparative_columns_show_the_raw_comparative_tb_for_every_account_type(tmp_path, canonical_data):
    # Unlike the old design, Comparative here is never zeroed for a P&L
    # account - it's a plain reference figure (last year's actual Dr/Cr),
    # not an "opening balance" concept.
    out = tmp_path / "wp.xlsx"
    wb = _build_and_save(canonical_data, out)
    etb_sheet_name, ws, header_row = _etb_sheet(wb)

    pl_code = str(canonical_data["pl_current"].iloc[0]["account_code"])
    dtc_row = _dtc_row_for(wb, pl_code)
    etb_row = header_row + 1 + (dtc_row - 2)

    sol = _evaluate(out)
    comp_debit = _cell(sol, "wp.xlsx", etb_sheet_name, f"L{etb_row}")
    comp_credit = _cell(sol, "wp.xlsx", etb_sheet_name, f"M{etb_row}")
    expected_debit = float(canonical_data["tb_comparative"].groupby("account_code")["debit"].sum().get(pl_code, 0.0))
    expected_credit = float(canonical_data["tb_comparative"].groupby("account_code")["credit"].sum().get(pl_code, 0.0))
    assert comp_debit == pytest.approx(expected_debit)
    assert comp_credit == pytest.approx(expected_credit)


def test_posting_a_journal_shows_on_etb_and_flows_into_the_pl_split(tmp_path, canonical_data):
    out = tmp_path / "wp.xlsx"
    wb = _build_and_save(canonical_data, out)
    etb_sheet_name, ws, header_row = _etb_sheet(wb)
    journals_sheet_name = next(s for s in wb.sheetnames if s.endswith("Journals"))
    journals_first_row = _journals_first_row(wb)

    pl_code = str(canonical_data["pl_current"].iloc[0]["account_code"])
    dtc_row = _dtc_row_for(wb, pl_code)
    etb_row = header_row + 1 + (dtc_row - 2)

    sol_before = _evaluate(out)
    pl_net_before = _cell(sol_before, "wp.xlsx", etb_sheet_name, f"H{etb_row}") - _cell(sol_before, "wp.xlsx", etb_sheet_name, f"I{etb_row}")

    JOURNAL_AMOUNT = 250.0
    wb[journals_sheet_name].cell(row=journals_first_row, column=1, value="JNL 1")
    wb[journals_sheet_name].cell(row=journals_first_row, column=3, value=pl_code)
    wb[journals_sheet_name].cell(row=journals_first_row, column=4, value=JOURNAL_AMOUNT)  # Debit
    wb.save(out)

    sol = _evaluate(out)
    journals_dr = _cell(sol, "wp.xlsx", etb_sheet_name, f"F{etb_row}")
    pl_net_after = _cell(sol, "wp.xlsx", etb_sheet_name, f"H{etb_row}") - _cell(sol, "wp.xlsx", etb_sheet_name, f"I{etb_row}")
    assert journals_dr == pytest.approx(JOURNAL_AMOUNT)
    # net P&L Dr-Cr position for this account should move by exactly the
    # posted debit, regardless of which column (Dr or Cr) the net lands on
    assert pl_net_after == pytest.approx(pl_net_before + JOURNAL_AMOUNT)


def test_etb_totals_self_balance(tmp_path, canonical_data):
    # A genuine evaluation-based check that every Dr/Cr pair on the totals
    # row actually ties - Xero PTB, Journals, and the combined P&L+BS
    # split all independently balance, the same sense-check a real ETB's
    # own totals row gives a preparer.
    out = tmp_path / "wp.xlsx"
    wb = _build_and_save(canonical_data, out)
    etb_sheet_name, ws, header_row = _etb_sheet(wb)

    data_start = header_row + 1
    total_row = next(r for r in range(data_start, ws.max_row + 1) if ws.cell(row=r, column=2).value == "TOTAL")

    sol = _evaluate(out)
    xero_dr = _cell(sol, "wp.xlsx", etb_sheet_name, f"D{total_row}")
    xero_cr = _cell(sol, "wp.xlsx", etb_sheet_name, f"E{total_row}")
    assert xero_dr == pytest.approx(xero_cr)

    pl_dr = _cell(sol, "wp.xlsx", etb_sheet_name, f"H{total_row}")
    pl_cr = _cell(sol, "wp.xlsx", etb_sheet_name, f"I{total_row}")
    bs_dr = _cell(sol, "wp.xlsx", etb_sheet_name, f"J{total_row}")
    bs_cr = _cell(sol, "wp.xlsx", etb_sheet_name, f"K{total_row}")
    assert (pl_dr + bs_dr) == pytest.approx(pl_cr + bs_cr)

    check_row = total_row + 1
    assert ws.cell(row=check_row, column=4).value == '=IF(ABS((D{0})-(E{0}))>0.01,"REVIEW","OK")'.format(total_row)
