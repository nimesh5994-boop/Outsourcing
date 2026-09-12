"""Unit tests for the P&L Notes section layout - see
excel_builder._write_pl_notes_sections, wired into build_recon_sheet.

Regression coverage for a real presentation gap found comparing our
output against a real, professionally-prepared working paper template: the
P&L Notes drill-down (pl_variance._pl_notes_detail) used to render as one
long flat table with the account name repeated on every row - useful as a
data export, but not something a reviewer or client would actually read as
"notes to the accounts". This checks the reformatted, per-account
"note" presentation instead: current/comparative/variance/driver in a
single header line, the contact-level breakdown underneath, one blank row
between notes.
"""
import pandas as pd
from openpyxl import Workbook

from app.excel_builder import build_recon_sheet
from app.recon import ReconResult


def _notes_detail() -> pd.DataFrame:
    return pd.DataFrame([
        {"Account": "400 - Advertising & Marketing", "Contact": "Harbour Heights Hotel",
         "Current Period Description": "Marketing", "Previous Period Description": "",
         "Current Year": 592.20, "Previous Year": 0.0, "Variance %": "", "Duplicate Name": "", "Main Variance Driver": ""},
        {"Account": "400 - Advertising & Marketing", "Contact": "TOTAL",
         "Current Period Description": "", "Previous Period Description": "",
         "Current Year": 762.20, "Previous Year": 0.0, "Variance %": 100.0, "Duplicate Name": "",
         "Main Variance Driver": "Higher - Harbour Heights Hotel (+£592.20)"},
        {"Account": "409 - Cleaning (house)", "Contact": "TOTAL",
         "Current Period Description": "", "Previous Period Description": "",
         "Current Year": 0.0, "Previous Year": 0.0, "Variance %": 0.0, "Duplicate Name": "", "Main Variance Driver": ""},
    ])


def _build_sheet():
    wb = Workbook()
    result = ReconResult(
        name="P&L variance review (current vs comparative)", status="review",
        message="1 P&L account(s) moved beyond materiality.",
        matched_detail=_notes_detail(),
        matched_detail_label="P&L Notes - supplier/customer analysis by nominal code",
    )
    build_recon_sheet(wb, "TEST", "Year ended 31 March 2026", "29", "29 PL Variance Review", result)
    return wb["29 PL Variance Review"]


def test_pl_notes_render_as_titled_sections_not_a_flat_table():
    ws = _build_sheet()
    values = [[c.value for c in row] for row in ws.iter_rows()]
    flat = [v for row in values for v in row if v is not None]

    header = next(v for v in flat if isinstance(v, str) and v.startswith("400 - Advertising & Marketing"))
    assert "Current year: £762.20" in header
    assert "Previous year: £0.00" in header
    assert "Variance: 100.0%" in header
    assert "Main driver: Higher - Harbour Heights Hotel (+£592.20)" in header

    # An account with no movement still gets a section, without a driver clause
    header_409 = next(v for v in flat if isinstance(v, str) and v.startswith("409 - Cleaning (house)"))
    assert "Main driver" not in header_409

    # The account name itself is no longer repeated on every contact row -
    # it only ever appears as the lead-in to the merged header line
    assert all(v == header or not str(v).startswith("400 - Advertising & Marketing") for v in flat)


def test_pl_notes_section_without_contact_detail_says_so():
    ws = _build_sheet()
    texts = [c.value for row in ws.iter_rows() for c in row if isinstance(c.value, str)]
    assert any("No individual supplier/customer detail" in t for t in texts)


def test_pl_notes_sections_are_separated_by_a_blank_row():
    ws = _build_sheet()
    rows_with_content = [r for r in range(1, ws.max_row + 1) if any(ws.cell(row=r, column=c).value is not None for c in range(1, 7))]
    # at least one gap (blank row) exists between the first and last content row,
    # proving sections aren't packed back-to-back with no separation
    assert rows_with_content[-1] - rows_with_content[0] + 1 > len(rows_with_content)
