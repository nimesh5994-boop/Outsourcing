"""Writes canonical DataFrames onto raw "DATA_*" sheets in the workbook, and
hands back cell-range references for each column - the foundation every
formula-linked schedule builds on. Every schedule cell should be a formula
referencing these ranges (or another schedule's own cells), not a
Python-computed literal, so the workbook recalculates like a manually-built
working paper and a reviewer can trace the formula chain.

A synthetic RowID column is added to nominal activity so a matrix-style
schedule (one row per transaction) can key a SUMIFS to exactly one source
row without relying on text criteria that might not be unique.
"""
from dataclasses import dataclass, field

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.tb_tieout import ADJUSTMENT_COLUMN_LETTER
from app.xlformulas import cell_ref, range_ref, sumifs_exact, quote

HEADER_FILL = PatternFill("solid", fgColor="D9E2F3")
HEADER_FONT = Font(bold=True)


@dataclass
class SheetRefs:
    """Column letter lookup + row bounds for one DATA_* sheet."""
    sheet_name: str
    columns: dict[str, str]  # canonical column -> Excel column letter
    first_row: int
    last_row: int  # first_row - 1 if the sheet has no data rows

    def col_range(self, column: str) -> str:
        return range_ref(self.sheet_name, self.columns[column], self.first_row, max(self.last_row, self.first_row))

    def is_empty(self) -> bool:
        return self.last_row < self.first_row


@dataclass
class DataRefs:
    tb_current: SheetRefs | None = None
    tb_comparative: SheetRefs | None = None
    nominal_current: SheetRefs | None = None
    aged_debtors: SheetRefs | None = None
    aged_creditors: SheetRefs | None = None
    pl_current: SheetRefs | None = None
    bs_current: SheetRefs | None = None
    fixed_asset_register: SheetRefs | None = None


def _cell_value(val):
    if isinstance(val, bool):
        return val
    if isinstance(val, pd.Timestamp):
        return val.date() if pd.notna(val) else None
    if isinstance(val, (int, float, str)):
        return val
    if val is None or pd.isna(val):
        return None
    return val


def _write_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame, balance_adjustment_cell=None) -> SheetRefs:
    """balance_adjustment_cell, when given, is a function (0-indexed row
    position -> a cell reference string) - the "balance" column is then
    written as a live formula (raw debit-credit, plus that cell) instead
    of a literal, so a live-linked TB Tie-Out sheet's Adjustment column
    can feed straight back into every schedule that reads this sheet's
    balance. debit/credit stay untouched literals either way."""
    ws: Worksheet = wb.create_sheet(sheet_name)
    columns = list(df.columns)
    col_letters = {col: get_column_letter(i + 1) for i, col in enumerate(columns)}

    for i, col in enumerate(columns):
        cell = ws.cell(row=1, column=i + 1, value=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    r = 2
    for row_i, (_, row) in enumerate(df.iterrows()):
        for i, col in enumerate(columns):
            if col == "balance" and balance_adjustment_cell is not None:
                debit_col, credit_col = col_letters["debit"], col_letters["credit"]
                ws.cell(row=r, column=i + 1, value=f"={debit_col}{r}-{credit_col}{r}+{balance_adjustment_cell(row_i)}")
            else:
                ws.cell(row=r, column=i + 1, value=_cell_value(row[col]))
        r += 1
    last_row = r - 1

    ws.sheet_state = "hidden"
    return SheetRefs(sheet_name=sheet_name, columns=col_letters, first_row=2, last_row=last_row)


def _write_derived_amount_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame, tb_current_refs, negate: bool) -> SheetRefs:
    """DATA_PL/DATA_BS: same rows xero_reports.derive_pl_bs_from_tb already
    split out (P&L vs Balance Sheet, by Account Type), but with "amount"
    written as a live lookup of that account's (possibly adjusted) balance
    on DATA_TB_Current, instead of the literal Python-computed figure - so
    the P&L/Balance Sheet statement sheets, which already read DATA_PL/
    DATA_BS through a live formula of their own, inherit a TB Tie-Out
    adjustment too, with no changes needed on their end. P&L rows negate
    the balance (income positive/expenses negative there, the opposite of
    DATA_TB_Current's debit-credit convention); B/S rows don't.
    Falls back to a plain literal (tb_current_refs is None - shouldn't
    happen in practice, since pl/bs are only ever derived FROM tb_current,
    but graceful either way) exactly like _write_sheet would have."""
    if tb_current_refs is None:
        return _write_sheet(wb, sheet_name, df)

    ws: Worksheet = wb.create_sheet(sheet_name)
    columns = list(df.columns)
    col_letters = {col: get_column_letter(i + 1) for i, col in enumerate(columns)}
    for i, col in enumerate(columns):
        cell = ws.cell(row=1, column=i + 1, value=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    code_range = tb_current_refs.col_range("account_code")
    balance_range = tb_current_refs.col_range("balance")
    r = 2
    for _, row in df.iterrows():
        for i, col in enumerate(columns):
            if col == "amount":
                lookup = sumifs_exact(balance_range, (code_range, quote(str(row["account_code"]))))
                ws.cell(row=r, column=i + 1, value=("=-" + lookup[1:]) if negate else lookup)
            else:
                ws.cell(row=r, column=i + 1, value=_cell_value(row[col]))
        r += 1
    last_row = r - 1

    ws.sheet_state = "hidden"
    return SheetRefs(sheet_name=sheet_name, columns=col_letters, first_row=2, last_row=last_row)


def with_row_ids(nominal_activity: pd.DataFrame) -> pd.DataFrame:
    """Adds the same synthetic RowID column write_data_sheets puts on
    DATA_Nominal, in the same row order - so Python-side code (e.g. the
    nominal matrix's formula-linked row grouping) can compute row_ids that
    line up exactly with the ones baked into the workbook."""
    df = nominal_activity.copy().reset_index(drop=True)
    df.insert(0, "row_id", range(1, len(df) + 1))
    return df


def write_data_sheets(wb: Workbook, data: dict, tb_adjustments_ref: tuple[str, int] | None = None) -> DataRefs:
    """tb_adjustments_ref, when given, is (sheet_name, first_data_row) for
    the TB Tie-Out sheet's Adjustment column (see tb_tieout.py's
    ADJUSTMENT_COLUMN_LETTER) - computed by the caller *before* that sheet
    is actually built (excel_builder._title_end_row lets it know the row
    without needing the sheet to exist yet), so DATA_TB_Current's own
    balance column can carry a live formula back to it: every downstream
    schedule that reads DATA_TB_Current's balance (directly, or via the
    P&L/B/S lookup below) sees whatever a preparer types there, with no
    changes needed in any of those schedules themselves. debit/credit stay
    untouched raw uploaded figures either way - only balance carries the
    adjustment - so the TB self-balance check and the TB Tie-Out's own
    "Reported Closing" column (which recomputes debit-credit directly,
    deliberately not through balance) keep showing the original, unadjusted
    upload for comparison."""
    refs = DataRefs()

    if data.get("tb_current") is not None and not data["tb_current"].empty:
        df = data["tb_current"][["account_code", "account_name", "account_type", "debit", "credit", "balance"]]
        adjustment_cell_for_row = None
        if tb_adjustments_ref is not None:
            adj_sheet, adj_first_row = tb_adjustments_ref
            adjustment_cell_for_row = lambda i: cell_ref(adj_sheet, f"{ADJUSTMENT_COLUMN_LETTER}{adj_first_row + i}")  # noqa: E731
        refs.tb_current = _write_sheet(wb, "DATA_TB_Current", df, balance_adjustment_cell=adjustment_cell_for_row)

    if data.get("tb_comparative") is not None and not data["tb_comparative"].empty:
        df = data["tb_comparative"][["account_code", "account_name", "account_type", "debit", "credit", "balance"]]
        refs.tb_comparative = _write_sheet(wb, "DATA_TB_Comparative", df)

    if data.get("nominal_current") is not None and not data["nominal_current"].empty:
        df = with_row_ids(data["nominal_current"])
        if "debit" in df.columns and "credit" in df.columns:
            df["net"] = df["debit"] - df["credit"]
        cols = ["row_id", "date", "account_code", "account_name", "reference", "description", "contact", "source_type", "debit", "credit", "net"]
        cols = [c for c in cols if c in df.columns]
        refs.nominal_current = _write_sheet(wb, "DATA_Nominal", df[cols])

    if data.get("aged_debtors") is not None and not data["aged_debtors"].empty:
        refs.aged_debtors = _write_sheet(wb, "DATA_AgedDebtors", data["aged_debtors"])

    if data.get("aged_creditors") is not None and not data["aged_creditors"].empty:
        refs.aged_creditors = _write_sheet(wb, "DATA_AgedCreditors", data["aged_creditors"])

    if data.get("pl_current") is not None and not data["pl_current"].empty:
        refs.pl_current = _write_derived_amount_sheet(wb, "DATA_PL", data["pl_current"], refs.tb_current, negate=True)

    if data.get("bs_current") is not None and not data["bs_current"].empty:
        refs.bs_current = _write_derived_amount_sheet(wb, "DATA_BS", data["bs_current"], refs.tb_current, negate=False)

    if data.get("fixed_asset_register") is not None and not data["fixed_asset_register"].empty:
        refs.fixed_asset_register = _write_sheet(wb, "DATA_FixedAssetRegister", data["fixed_asset_register"])

    return refs
