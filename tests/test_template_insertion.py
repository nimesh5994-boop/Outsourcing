"""Tests for generating into a copy of a practice's real uploaded template
file, instead of the system's own generic layout - verifies the template's
own sheets are never touched, schedules land where the config says, the
enabled/disabled toggle is honoured, and numbering respects start_at."""
import copy

import openpyxl
import pytest

from app import control_accounts, corporation_tax, fixed_assets, nominal_matrix, recon
from app.excel_builder import build_workbook_into_template
from app.storage import DEFAULT_TEMPLATE_CONFIG


def _make_template(tmp_path, sheet_names):
    wb = openpyxl.Workbook()
    wb.active.title = sheet_names[0]
    wb.active["A1"] = f"original content - {sheet_names[0]}"
    for name in sheet_names[1:]:
        ws = wb.create_sheet(name)
        ws["A1"] = f"original content - {name}"
    path = tmp_path / "practice_template.xlsx"
    wb.save(path)
    return path


@pytest.fixture
def generation_inputs(canonical_data):
    results = recon.run_all_recons(canonical_data)
    ca_results = control_accounts.build_all_rollforwards(
        canonical_data["tb_current"], canonical_data["tb_comparative"], canonical_data["nominal_current"],
        canonical_data["aged_debtors"], canonical_data["aged_creditors"],
    )
    mx_results = nominal_matrix.build_all_matrices(canonical_data["tb_current"], canonical_data["nominal_current"])
    ct_computation = corporation_tax.compute(accounting_profit=float(canonical_data["pl_current"]["amount"].sum()))
    fixed_asset_result = fixed_assets.category_level_rollforward(
        canonical_data["tb_current"], canonical_data["tb_comparative"], canonical_data["nominal_current"]
    )
    return {
        "data": canonical_data, "results": results, "control_account_results": ca_results,
        "matrix_results": mx_results, "ct_computation": ct_computation, "fixed_asset_result": fixed_asset_result,
    }


def test_template_sheets_are_preserved_untouched(tmp_path, generation_inputs):
    template_path = _make_template(tmp_path, ["Cover Page", "Firm Notes"])
    config = copy.deepcopy(DEFAULT_TEMPLATE_CONFIG)

    wb = build_workbook_into_template(
        template_path, config, "Brightwell Landscaping Supplies Limited",
        "Year ended 31 December 2025", "Year ended 31 December 2024", **generation_inputs,
    )

    assert "Cover Page" in wb.sheetnames
    assert "Firm Notes" in wb.sheetnames
    assert wb["Cover Page"]["A1"].value == "original content - Cover Page"
    assert wb["Firm Notes"]["A1"].value == "original content - Firm Notes"
    # generated content added on top, not just the two original sheets
    assert len(wb.sheetnames) > 2


def test_index_sheet_does_not_overwrite_templates_own_first_sheet(tmp_path, generation_inputs):
    template_path = _make_template(tmp_path, ["Cover Page"])
    config = copy.deepcopy(DEFAULT_TEMPLATE_CONFIG)

    wb = build_workbook_into_template(
        template_path, config, "Brightwell Landscaping Supplies Limited",
        "Year ended 31 December 2025", "Year ended 31 December 2024", **generation_inputs,
    )

    assert wb["Cover Page"]["A1"].value == "original content - Cover Page"
    assert "Index" in wb.sheetnames
    # default (no insert_after_sheet configured): index goes first
    assert wb.sheetnames.index("Index") < wb.sheetnames.index("Cover Page")


def test_disabled_schedule_is_skipped(tmp_path, generation_inputs):
    template_path = _make_template(tmp_path, ["Cover Page"])
    config = copy.deepcopy(DEFAULT_TEMPLATE_CONFIG)
    config["schedules"]["corporation_tax"]["enabled"] = False

    wb = build_workbook_into_template(
        template_path, config, "Brightwell Landscaping Supplies Limited",
        "Year ended 31 December 2025", "Year ended 31 December 2024", **generation_inputs,
    )

    assert not any("Corporation Tax" in name for name in wb.sheetnames)
    assert any("TB Lead Schedule" in name for name in wb.sheetnames)  # sibling schedules unaffected


def test_insert_after_sheet_positions_the_schedule(tmp_path, generation_inputs):
    template_path = _make_template(tmp_path, ["Cover Page", "Anchor Sheet", "Back Matter"])
    config = copy.deepcopy(DEFAULT_TEMPLATE_CONFIG)
    config["schedules"]["tb_lead_schedule"]["insert_after_sheet"] = "Anchor Sheet"

    wb = build_workbook_into_template(
        template_path, config, "Brightwell Landscaping Supplies Limited",
        "Year ended 31 December 2025", "Year ended 31 December 2024", **generation_inputs,
    )

    tb_sheet = next(n for n in wb.sheetnames if "TB Lead Schedule" in n)
    anchor_idx = wb.sheetnames.index("Anchor Sheet")
    back_matter_idx = wb.sheetnames.index("Back Matter")
    tb_idx = wb.sheetnames.index(tb_sheet)
    assert anchor_idx < tb_idx < back_matter_idx


def test_numbering_start_at_is_respected(tmp_path, generation_inputs):
    template_path = _make_template(tmp_path, ["Cover Page"])
    config = copy.deepcopy(DEFAULT_TEMPLATE_CONFIG)
    config["numbering"]["start_at"] = 21

    wb = build_workbook_into_template(
        template_path, config, "Brightwell Landscaping Supplies Limited",
        "Year ended 31 December 2025", "Year ended 31 December 2024", **generation_inputs,
    )

    assert any(n.startswith("21 ") for n in wb.sheetnames)
    assert not any(n.startswith("1 ") for n in wb.sheetnames)


def test_sheet_name_collision_with_template_is_handled_not_fatal(tmp_path, generation_inputs):
    # the template already has its own sheet literally named "Index" -
    # openpyxl auto-suffixes ours rather than erroring or overwriting
    template_path = _make_template(tmp_path, ["Index", "Cover Page"])
    config = copy.deepcopy(DEFAULT_TEMPLATE_CONFIG)

    wb = build_workbook_into_template(
        template_path, config, "Brightwell Landscaping Supplies Limited",
        "Year ended 31 December 2025", "Year ended 31 December 2024", **generation_inputs,
    )

    assert wb["Index"]["A1"].value == "original content - Index"
    assert any(n != "Index" and "Index" in n for n in wb.sheetnames)


def test_formula_linked_schedules_still_reference_data_sheets_in_template_mode(tmp_path, generation_inputs):
    # the formula-linked engine must work the same way when the base
    # workbook is a loaded template, not just a fresh Workbook()
    template_path = _make_template(tmp_path, ["Cover Page"])
    config = copy.deepcopy(DEFAULT_TEMPLATE_CONFIG)

    wb = build_workbook_into_template(
        template_path, config, "Brightwell Landscaping Supplies Limited",
        "Year ended 31 December 2025", "Year ended 31 December 2024", **generation_inputs,
    )

    assert "DATA_TB_Current" in wb.sheetnames
    tb_sheet_name = next(n for n in wb.sheetnames if "TB Lead Schedule" in n)
    ws = wb[tb_sheet_name]
    formula_cells = [c.value for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("=")]
    assert formula_cells, "expected live formulas on the TB Lead Schedule sheet even in template mode"
