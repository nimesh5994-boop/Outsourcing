"""Builds the final Excel working paper pack from canonical data + recon
results, in the house style observed in the firm's real working paper files:
a CLIENT NAME / PERIOD / SCHEDULE TITLE header block on every sheet, and a
numbered index of schedules with their status and any queries raised.
"""
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.control_accounts import ControlAccountResult
from app.nominal_matrix import MatrixResult
from app.recon import ReconResult

NAVY = "1F3864"
GREEN = "C6EFCE"
GREEN_TEXT = "006100"
AMBER = "FFEB9C"
AMBER_TEXT = "9C6500"
RED = "FFC7CE"
RED_TEXT = "9C0006"
GREY = "F2F2F2"

HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
CLIENT_FONT = Font(name="Calibri", size=13, bold=True, color=NAVY)
PERIOD_FONT = Font(name="Calibri", size=11, color="595959")
SCHEDULE_FONT = Font(name="Calibri", size=12, bold=True, color="000000")
SUBTITLE_FONT = Font(name="Calibri", size=11, italic=True, color="595959")
BOLD = Font(name="Calibri", size=11, bold=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CURRENCY_FMT = '#,##0.00;[RED](#,##0.00)'


def _style_header_row(ws: Worksheet, row: int, ncols: int):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def _autosize(ws: Worksheet, df: pd.DataFrame, start_col: int = 1):
    for i, col in enumerate(df.columns):
        width = max(12, min(45, int(df[col].astype(str).str.len().max() if len(df) else 0), len(str(col)) + 2))
        width = max(width, len(str(col)) + 2)
        ws.column_dimensions[get_column_letter(start_col + i)].width = width


def _write_title(ws: Worksheet, client_name: str, period_label: str, schedule_title: str, ref: str = "") -> int:
    """House-style header block: CLIENT NAME / PERIOD / SCHEDULE TITLE.
    Returns the first free row below the header block."""
    ws["A1"] = client_name.upper()
    ws["A1"].font = CLIENT_FONT
    ws["A2"] = period_label.upper()
    ws["A2"].font = PERIOD_FONT
    ws["A3"] = f"{ref + '  ' if ref else ''}{schedule_title}"
    ws["A3"].font = SCHEDULE_FONT
    return 5


def _status_fill(status: str) -> PatternFill:
    return {
        "ok": PatternFill("solid", fgColor=GREEN),
        "review": PatternFill("solid", fgColor=AMBER),
        "error": PatternFill("solid", fgColor=RED),
        "n/a": PatternFill("solid", fgColor=GREY),
    }.get(status, PatternFill("solid", fgColor=GREY))


def _status_font(status: str) -> Font:
    color = {"ok": GREEN_TEXT, "review": AMBER_TEXT, "error": RED_TEXT, "n/a": "595959"}.get(status, "000000")
    return Font(bold=True, color=color)


def _write_dataframe(ws: Worksheet, df: pd.DataFrame, start_row: int, start_col: int = 1) -> int:
    """Writes header + rows starting at start_row, returns the next free row."""
    for j, col in enumerate(df.columns):
        ws.cell(row=start_row, column=start_col + j, value=str(col))
    _style_header_row(ws, start_row, len(df.columns))
    r = start_row + 1
    for _, row in df.iterrows():
        for j, col in enumerate(df.columns):
            val = row[col]
            if isinstance(val, pd.Timestamp):
                val = val.date() if not pd.isna(val) else ""
            cell = ws.cell(row=r, column=start_col + j, value=val)
            cell.border = BORDER
            if isinstance(val, (int, float)):
                cell.number_format = CURRENCY_FMT
        r += 1
    _autosize(ws, df, start_col)
    return r + 1


class _RefCounter:
    def __init__(self):
        self.n = 0

    def next(self) -> str:
        self.n += 1
        return str(self.n)


def build_index_sheet(wb: Workbook, client_name: str, current_label: str, comparative_label: str, entries: list[dict]):
    ws = wb.active
    ws.title = "Index"
    ws["A1"] = client_name.upper()
    ws["A1"].font = CLIENT_FONT
    ws["A2"] = f"WORKING PAPERS - {current_label.upper()}"
    ws["A2"].font = PERIOD_FONT
    ws["A3"] = f"Comparative period: {comparative_label or 'none'}"
    ws["A3"].font = SUBTITLE_FONT
    ws["A4"] = f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}"
    ws["A4"].font = SUBTITLE_FONT

    headers = ["Ref", "Schedule", "Status", "Comment"]
    start = 6
    for j, h in enumerate(headers):
        ws.cell(row=start, column=1 + j, value=h)
    _style_header_row(ws, start, len(headers))

    r = start + 1
    for e in entries:
        ws.cell(row=r, column=1, value=e["ref"]).border = BORDER
        ws.cell(row=r, column=2, value=e["title"]).border = BORDER
        status_cell = ws.cell(row=r, column=3, value=e["status"].upper())
        status_cell.fill = _status_fill(e["status"])
        status_cell.font = _status_font(e["status"])
        status_cell.alignment = Alignment(horizontal="center")
        status_cell.border = BORDER
        comment_cell = ws.cell(row=r, column=4, value=e["message"])
        comment_cell.border = BORDER
        comment_cell.alignment = Alignment(wrap_text=True)
        r += 1

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 90
    ws.freeze_panes = "A7"


def build_tb_lead_schedule(wb: Workbook, client_name: str, current_label: str, ref: str, variance_detail: pd.DataFrame):
    ws = wb.create_sheet(f"{ref} TB Lead Schedule"[:31])
    row = _write_title(ws, client_name, current_label, "TRIAL BALANCE LEAD SCHEDULE", ref)
    if variance_detail is None or variance_detail.empty:
        ws.cell(row=row, column=1, value="No trial balance data available.")
        return

    df = variance_detail.copy()
    df["Reviewed / reallocation required?"] = ""
    df = df.rename(columns={
        "account_code": "Account Code", "account_name": "Account Name",
        "current_year": "Current Year", "comparative_year": "Comparative Year",
        "variance_amount": "Variance (£)", "variance_pct": "Variance (%)", "flag": "Flagged?",
    })
    df["Variance (%)"] = df["Variance (%)"].apply(lambda x: round(x * 100, 1))
    next_row = _write_dataframe(ws, df, start_row=row)

    flag_col = list(df.columns).index("Flagged?") + 1
    for r in range(row + 1, next_row - 1):
        if ws.cell(row=r, column=flag_col).value is True:
            for c in range(1, len(df.columns) + 1):
                ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=AMBER)
    ws.freeze_panes = f"A{row + 1}"


def build_statement_sheet(wb: Workbook, client_name: str, period_label: str, ref: str, sheet_name: str, title: str, df: pd.DataFrame):
    ws = wb.create_sheet(sheet_name[:31])
    row = _write_title(ws, client_name, period_label, title, ref)
    if df is None or df.empty:
        ws.cell(row=row, column=1, value="No data uploaded for this statement.")
        return
    display = df.rename(columns={"account_code": "Account Code", "account_name": "Account Name", "category": "Category", "amount": "Amount"})
    _write_dataframe(ws, display, start_row=row)
    ws.freeze_panes = f"A{row + 1}"


def build_recon_sheet(wb: Workbook, client_name: str, period_label: str, ref: str, sheet_name: str, result: ReconResult):
    ws = wb.create_sheet(sheet_name[:31])
    row = _write_title(ws, client_name, period_label, result.name, ref)
    status_cell = ws.cell(row=row, column=1, value=f"Status: {result.status.upper()} - {result.message}")
    status_cell.font = _status_font(result.status)
    status_cell.fill = _status_fill(result.status)
    status_cell.alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.row_dimensions[row].height = 30

    detail_row = row + 2
    if result.detail is None or result.detail.empty:
        ws.cell(row=detail_row, column=1, value="No supporting detail available.")
        return
    _write_dataframe(ws, result.detail, start_row=detail_row)
    ws.freeze_panes = f"A{detail_row + 1}"


def build_control_account_sheet(wb: Workbook, client_name: str, current_label: str, ref: str, result: ControlAccountResult):
    sheet_name = f"{ref} {result.account_name}"[:31]
    ws = wb.create_sheet(sheet_name)
    suffix = "" if "control account" in result.account_name.lower() else " CONTROL ACCOUNT"
    title = f"{result.account_name.upper()}{suffix} (nominal {result.account_code})"
    row = _write_title(ws, client_name, current_label, title, ref)

    status_cell = ws.cell(row=row, column=1, value=f"Status: {result.status.upper()} - {result.message}")
    status_cell.font = _status_font(result.status)
    status_cell.fill = _status_fill(result.status)
    status_cell.alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    ws.row_dimensions[row].height = 30

    table_row = row + 2
    if result.schedule.empty and result.breakdown.empty:
        ws.cell(row=table_row, column=1, value="No data available.")
        return

    next_row = table_row
    if not result.schedule.empty:
        next_row = _write_dataframe(ws, result.schedule, start_row=table_row)
        diff_row = next_row - 2
        if str(ws.cell(row=diff_row, column=1).value or "").startswith("DIFFERENCE"):
            for c in range(1, 5):
                ws.cell(row=diff_row, column=c).font = BOLD
                if result.status != "ok":
                    ws.cell(row=diff_row, column=c).fill = PatternFill("solid", fgColor=AMBER)

    if not result.breakdown.empty:
        header_row = next_row
        ws.cell(row=header_row, column=1, value=f"BREAKDOWN OF CLOSING BALANCE ({result.breakdown_label})").font = SCHEDULE_FONT
        breakdown_start = header_row + 1
        end_row = _write_dataframe(ws, result.breakdown, start_row=breakdown_start)
        # bold + flag the three summary rows appended at the bottom (total / TB balance / unexplained diff)
        for offset in range(1, 4):
            r = end_row - 1 - offset
            for c in range(1, 3):
                ws.cell(row=r, column=c).font = BOLD
        diff_row = end_row - 2
        if result.status != "ok" and str(ws.cell(row=diff_row, column=1).value or "").startswith("UNEXPLAINED"):
            for c in range(1, 3):
                ws.cell(row=diff_row, column=c).fill = PatternFill("solid", fgColor=AMBER)

    ws.freeze_panes = f"A{table_row + 1}"


def build_matrix_sheet(wb: Workbook, client_name: str, current_label: str, ref: str, result: MatrixResult):
    sheet_name = f"{ref} {result.account_name} Analysis"[:31]
    ws = wb.create_sheet(sheet_name)
    title = f"NOMINAL ACTIVITY ANALYSIS - {result.account_name} ({result.account_code})"
    row = _write_title(ws, client_name, current_label, title, ref)

    status_cell = ws.cell(row=row, column=1, value=f"Status: {result.status.upper()} - {result.message}")
    status_cell.font = _status_font(result.status)
    status_cell.fill = _status_fill(result.status)
    status_cell.alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.row_dimensions[row].height = 30

    table_row = row + 2
    if result.matrix is None or result.matrix.empty:
        ws.cell(row=table_row, column=1, value="No data available.")
        return
    df = result.matrix.rename(columns={"date": "Date", "reference": "Reference", "description": "Description", "contact": "Contact"})
    _write_dataframe(ws, df, start_row=table_row)
    ws.freeze_panes = f"A{table_row + 1}"


def build_points_forward_sheet(wb: Workbook, client_name: str, period_label: str, ref: str):
    ws = wb.create_sheet(f"{ref} Points Forward"[:31])
    row = _write_title(ws, client_name, period_label, "POINTS FORWARD / OPEN QUERIES", ref)
    headers = ["#", "Area", "Query", "Raised by", "Date raised", "Response / Resolution", "Status"]
    for j, h in enumerate(headers):
        ws.cell(row=row, column=1 + j, value=h)
    _style_header_row(ws, row, len(headers))
    for r in range(row + 1, row + 21):
        ws.cell(row=r, column=1, value=r - row)
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).border = BORDER
    widths = [4, 18, 45, 14, 14, 45, 14]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(1 + i)].width = w
    ws.freeze_panes = f"A{row + 1}"


def build_workbook(
    client_name: str,
    current_label: str,
    comparative_label: str,
    data: dict,
    results: list[ReconResult],
    control_account_results: list[ControlAccountResult] | None = None,
    matrix_results: list[MatrixResult] | None = None,
) -> Workbook:
    control_account_results = control_account_results or []
    matrix_results = matrix_results or []

    ref = _RefCounter()
    entries = []
    wb = Workbook()

    tb_ref = ref.next()
    entries.append({"ref": tb_ref, "title": "TB Lead Schedule", "status": "ok" if data.get("tb_current") is not None and not data["tb_current"].empty else "n/a", "message": "Current vs comparative, variance flagged"})

    pl_ref, bs_ref = ref.next(), ref.next()
    entries.append({"ref": pl_ref, "title": "Profit & Loss", "status": "ok" if data.get("pl_current") is not None and not data["pl_current"].empty else "n/a", "message": ""})
    entries.append({"ref": bs_ref, "title": "Balance Sheet", "status": "ok" if data.get("bs_current") is not None and not data["bs_current"].empty else "n/a", "message": ""})

    name_map = {
        "TB self-balance check": "TB Balance Check",
        "TB to P&L/B&S tie-out (current year)": "TB to P&L-BS Tie",
        "Current vs comparative variance analysis": None,  # feeds the lead schedule, no separate tab
        "Debtors control account reconciliation": "Debtors Recon",
        "Creditors control account reconciliation": "Creditors Recon",
        "Bank reconciliation": "Bank Recon",
        "VAT return cross-check": "VAT Recon",
        "Nominal activity review (reallocation candidates)": "Nominal Review",
    }
    recon_refs = {}
    for res in results:
        sheet_name = name_map.get(res.name, res.name[:31])
        if sheet_name is None:
            continue
        recon_refs[res.name] = ref.next()
        entries.append({"ref": recon_refs[res.name], "title": res.name, "status": res.status, "message": res.message})

    ca_refs = {}
    for r in control_account_results:
        ca_refs[r.account_code] = ref.next()
        entries.append({"ref": ca_refs[r.account_code], "title": f"{r.account_name} control account", "status": r.status, "message": r.message})

    mx_refs = {}
    for r in matrix_results:
        mx_refs[r.account_code] = ref.next()
        entries.append({"ref": mx_refs[r.account_code], "title": f"{r.account_name} nominal analysis", "status": r.status, "message": r.message})

    pf_ref = ref.next()
    entries.append({"ref": pf_ref, "title": "Points Forward / Open Queries", "status": "n/a", "message": "To be completed by preparer/reviewer"})

    build_index_sheet(wb, client_name, current_label, comparative_label, entries)

    variance_result = next((r for r in results if r.name == "Current vs comparative variance analysis"), None)
    build_tb_lead_schedule(wb, client_name, current_label, tb_ref, variance_result.detail if variance_result else pd.DataFrame())

    build_statement_sheet(wb, client_name, current_label, pl_ref, f"{pl_ref} P&L", "PROFIT & LOSS", data.get("pl_current"))
    build_statement_sheet(wb, client_name, current_label, bs_ref, f"{bs_ref} Balance Sheet", "BALANCE SHEET", data.get("bs_current"))

    for res in results:
        sheet_name = name_map.get(res.name, res.name[:31])
        if sheet_name is None:
            continue
        r = recon_refs[res.name]
        build_recon_sheet(wb, client_name, current_label, r, f"{r} {sheet_name}", res)

    for r in control_account_results:
        build_control_account_sheet(wb, client_name, current_label, ca_refs[r.account_code], r)

    for r in matrix_results:
        build_matrix_sheet(wb, client_name, current_label, mx_refs[r.account_code], r)

    build_points_forward_sheet(wb, client_name, current_label, pf_ref)
    return wb
