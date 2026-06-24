"""Tests for the Excel mapping-pattern (template) workflow.

Covers: copying one worksheet's column pattern to multiple mapped sheets,
per-sheet override after copying, the final write plan using the correct
assigned sheet per Customer Name, and the guarantee that groups are NOT all
written into the template worksheet.

All sheet/column/customer names below are arbitrary values supplied at test
time; none of them appear in application code.
"""
import datetime as dt
import io

import openpyxl
from openpyxl.utils import column_index_from_string

from src import excel_reader, mapping
from src.aggregator import aggregate
from src.excel_writer import apply_writes
from src.mapping import (
    AppConfig, OutputType, SheetMapping, SourceMapping, TargetMode,
    WriteAction, WriteRules, apply_pattern_to_sheets, copy_sheet_mapping,
)
from src.parser import parse_text_lines
from src.validator import (
    MapStatus, Status, build_plan, configured_mapped_sheets,
    plan_to_write_ops, sheet_mapping_status,
)


def _workbook_three_sheets():
    """Three branch sheets sharing the SAME layout: A=Date, B=Amount, C=Return,
    plus an unrelated sheet with a different layout."""
    wb = openpyxl.Workbook()
    first = True
    for nm in ["SH_ONE", "SH_TWO", "SH_THREE"]:
        ws = wb.active if first else wb.create_sheet(nm)
        if first:
            ws.title = nm
            first = False
        ws.append(["Date", "Amount", "Return"])
        for i, day in enumerate((15, 16, 17), start=2):
            ws.cell(row=i, column=1, value=dt.date(2026, 5, day))
    # A sheet with a DIFFERENT structure to prove overrides matter.
    odd = wb.create_sheet("SH_ODD")
    odd.append(["When", "Val", "Ded"])
    odd.cell(row=2, column=1, value=dt.date(2026, 5, 15))
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def _template_mapping():
    return SheetMapping(
        sheet_name="SH_ONE", header_row=1,
        date_target_mode=TargetMode.COLUMN_LETTER, date_column="A",
        amount_target_mode=TargetMode.COLUMN_LETTER, amount_column="B",
        return_target_mode=TargetMode.COLUMN_LETTER, return_column="C",
        date_formats=["%d/%m/%Y"],
    )


def _config(group_to_sheet):
    cfg = AppConfig(
        source=SourceMapping(date_field="Inv Date", amount_field="Amount",
                             return_field="Ret", sum_duplicate_dates=True),
        group_to_sheet=dict(group_to_sheet),
        sheets={"SH_ONE": _template_mapping()},
        write_rules=WriteRules(write_action=WriteAction.REPLACE,
                               output_type=OutputType.NUMERIC, aggregation_keys=["group", "date"]),
    )
    return cfg


def test_copy_pattern_to_multiple_sheets():
    cfg = _config({"GROUP A": "SH_ONE", "GROUP B": "SH_TWO", "GROUP C": "SH_THREE"})
    written = apply_pattern_to_sheets(cfg, "SH_ONE", cfg.group_to_sheet.values())
    assert set(written) == {"SH_TWO", "SH_THREE"}          # template itself not rewritten
    # Each target now has the same column pattern, but its OWN sheet_name.
    for nm in ("SH_TWO", "SH_THREE"):
        sm = cfg.sheets[nm]
        assert sm.sheet_name == nm
        assert sm.amount_column == "B"
        assert sm.date_column == "A"
        assert sm.amount_target_mode == TargetMode.COLUMN_LETTER
        assert sm.date_formats == ["%d/%m/%Y"]


def test_cell_reference_pattern_downgrades_to_column_letter_on_copy():
    tmpl = _template_mapping()
    tmpl.amount_target_mode = TargetMode.CELL_REFERENCE
    tmpl.amount_cell = "B10"
    copy = copy_sheet_mapping(tmpl, "OTHER")
    # Exact cell refs are sheet-specific; copies fall back to column letter.
    assert copy.amount_cell == ""
    assert copy.amount_target_mode == TargetMode.COLUMN_LETTER


def test_per_sheet_override_after_copying():
    cfg = _config({"GROUP A": "SH_ONE", "GROUP B": "SH_TWO", "GROUP C": "SH_ODD"})
    apply_pattern_to_sheets(cfg, "SH_ONE", cfg.group_to_sheet.values())
    # SH_ODD got the pattern (A/B/C) but its real columns are When/Val/Ded.
    # Override it individually.
    cfg.sheets["SH_ODD"] = SheetMapping(
        sheet_name="SH_ODD", header_row=1,
        date_target_mode=TargetMode.HEADER_NAME, date_column="When",
        amount_target_mode=TargetMode.HEADER_NAME, amount_column="Val",
    )
    # Other sheets keep the pattern; the override is isolated.
    assert cfg.sheets["SH_TWO"].amount_column == "B"
    assert cfg.sheets["SH_ODD"].amount_column == "Val"
    assert cfg.sheets["SH_ODD"].date_column == "When"


def test_mapping_status_reports_per_group():
    data = _workbook_three_sheets()
    cfg = _config({"GROUP A": "SH_ONE", "GROUP B": "SH_TWO", "GROUP C": "SH_THREE",
                   "GROUP D": "NO_SUCH_SHEET"})
    apply_pattern_to_sheets(cfg, "SH_ONE", cfg.group_to_sheet.values())
    wb = excel_reader.open_workbook(data)
    rows = {r["customer_name"]: r for r in sheet_mapping_status(wb, cfg)}
    assert rows["GROUP A"]["status"] == MapStatus.READY
    assert rows["GROUP B"]["status"] == MapStatus.READY
    assert rows["GROUP D"]["status"] == MapStatus.SHEET_MISSING
    assert rows["GROUP B"]["assigned_sheet"] == "SH_TWO"


def test_configured_mapped_sheets_warning_condition():
    cfg = _config({"GROUP A": "SH_ONE", "GROUP B": "SH_TWO", "GROUP C": "SH_THREE"})
    # Only the template is configured initially -> warning condition true.
    assert configured_mapped_sheets(cfg) == ["SH_ONE"]
    apply_pattern_to_sheets(cfg, "SH_ONE", cfg.group_to_sheet.values())
    assert set(configured_mapped_sheets(cfg)) == {"SH_ONE", "SH_TWO", "SH_THREE"}


def _records_two_groups():
    text = """\
Customer Name: GROUP A
A1   15/05/2026   100.00   0.00
Customer Name: GROUP B
B1   16/05/2026   200.00   0.00
Customer Name: GROUP C
C1   17/05/2026   300.00   0.00
"""
    return parse_text_lines([text], group_label="Customer Name",
                            field_names=["Inv No", "Inv Date", "Amount", "Ret"])


def test_final_plan_uses_each_groups_assigned_sheet():
    data = _workbook_three_sheets()
    cfg = _config({"GROUP A": "SH_ONE", "GROUP B": "SH_TWO", "GROUP C": "SH_THREE"})
    apply_pattern_to_sheets(cfg, "SH_ONE", cfg.group_to_sheet.values())
    rows = aggregate(_records_two_groups(), cfg.source, aggregation_keys=["group", "date"])
    wb = excel_reader.open_workbook(data)
    report = build_plan(wb, rows, cfg)
    by_group = {i.group: i for i in report.items}
    assert by_group["GROUP A"].sheet == "SH_ONE"
    assert by_group["GROUP B"].sheet == "SH_TWO"
    assert by_group["GROUP C"].sheet == "SH_THREE"
    assert all(i.status == Status.READY for i in report.items)


def test_groups_not_all_written_to_template_sheet():
    data = _workbook_three_sheets()
    cfg = _config({"GROUP A": "SH_ONE", "GROUP B": "SH_TWO", "GROUP C": "SH_THREE"})
    apply_pattern_to_sheets(cfg, "SH_ONE", cfg.group_to_sheet.values())
    rows = aggregate(_records_two_groups(), cfg.source, aggregation_keys=["group", "date"])
    wb = excel_reader.open_workbook(data)
    report = build_plan(wb, rows, cfg)
    out, _ = apply_writes(data, plan_to_write_ops(report, cfg))
    out_wb = openpyxl.load_workbook(io.BytesIO(out))
    bcol = column_index_from_string("B")
    # Each value landed in ITS OWN sheet, NOT all piled into the template.
    assert out_wb["SH_ONE"].cell(row=2, column=bcol).value == 100.0  # 15/05 GROUP A
    assert out_wb["SH_TWO"].cell(row=3, column=bcol).value == 200.0  # 16/05 GROUP B
    assert out_wb["SH_THREE"].cell(row=4, column=bcol).value == 300.0  # 17/05 GROUP C
    # The template sheet must NOT contain GROUP B / GROUP C amounts.
    one_values = [out_wb["SH_ONE"].cell(row=r, column=bcol).value for r in range(2, 5)]
    assert 200.0 not in one_values
    assert 300.0 not in one_values
