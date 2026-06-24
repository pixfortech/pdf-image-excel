"""Tests proving Excel date cells are read by their REAL stored value, not by
their display formatting.

A cell formatted ``d-mmm`` shows e.g. ``04-Jan`` but stores a real
``datetime(2026, 1, 4)``.  openpyxl returns the real datetime, so matching must
use the true year — never infer the year from the displayed text.
"""
import datetime as dt
import io

import openpyxl
from openpyxl.utils import column_index_from_string

from src import excel_reader
from src.aggregator import aggregate
from src.mapping import (
    AppConfig, SheetMapping, SourceMapping, TargetMode, WriteAction, WriteRules,
)
from src.parser import parse_text_lines
from src.validator import Status, build_plan, excel_date_debug_rows


def _wb_with_format(number_format, the_date=dt.date(2026, 1, 4), n_extra=0):
    """Workbook whose date cell stores a real date but displays via the given
    custom number format. ``n_extra`` pads rows BEFORE the date to test full
    column scanning beyond the preview window."""
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "S1"
    ws.append(["Date", "Amount"])
    for _ in range(n_extra):
        ws.append([None, None])
    c = ws.cell(row=ws.max_row + 1, column=1, value=dt.datetime(the_date.year, the_date.month, the_date.day))
    c.number_format = number_format
    ws.cell(row=c.row, column=2, value=None)
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue(), c.row


def test_d_mmm_display_stores_real_year():
    data, row = _wb_with_format("d-mmm")
    wb = excel_reader.open_workbook(data)
    cell = wb["S1"].cell(row=row, column=1)
    # Displays "04-Jan" but the real value carries the year 2026.
    assert cell.number_format == "d-mmm"
    assert cell.value == dt.datetime(2026, 1, 4)
    found = excel_reader.find_row_by_date(wb, "S1", 1, dt.date(2026, 1, 4), header_row=1)
    assert found == row


def test_dd_mmm_display_matches_pdf_date():
    data, row = _wb_with_format("dd-mmm")
    wb = excel_reader.open_workbook(data)
    found = excel_reader.find_row_by_date(wb, "S1", 1, dt.date(2026, 1, 4), header_row=1)
    assert found == row


def test_full_column_scan_match_far_below_preview():
    # Date is 50 rows down, well past any 10-row preview window.
    data, row = _wb_with_format("d-mmm", n_extra=50)
    assert row > 50
    wb = excel_reader.open_workbook(data)
    found = excel_reader.find_row_by_date(wb, "S1", 1, dt.date(2026, 1, 4), header_row=1)
    assert found == row


def _config():
    return AppConfig(
        source=SourceMapping(date_field="Inv Date", amount_field="Amount",
                             sum_duplicate_dates=True, date_interpretation="dmy"),
        group_to_sheet={"GRP": "S1"},
        sheets={"S1": SheetMapping(sheet_name="S1", header_row=1,
                                   date_target_mode=TargetMode.COLUMN_LETTER, date_column="A",
                                   amount_target_mode=TargetMode.COLUMN_LETTER, amount_column="B")},
        write_rules=WriteRules(write_action=WriteAction.REPLACE, aggregation_keys=["group", "date"]),
    )


def test_pdf_date_matches_d_mmm_row_end_to_end():
    data, row = _wb_with_format("d-mmm")
    records = parse_text_lines(["Customer Name: GRP\nINV1   04/01/2026   500.00"],
                               group_label="Customer Name",
                               field_names=["Inv No", "Inv Date", "Amount"])
    cfg = _config()
    rows = aggregate(records, cfg.source, aggregation_keys=["group", "date"])
    report = build_plan(excel_reader.open_workbook(data), rows, cfg)
    item = report.items[0]
    assert item.status == Status.READY
    assert item.matched_row == row
    assert item.pdf_date_normalised == "2026-01-04"
    assert item.excel_date_normalised == "2026-01-04"


def test_in_range_but_missing_is_flagged_error():
    # Excel has 2026-01-01 and 2026-12-31; PDF date 2026-06-15 is in range but
    # not present -> must be flagged as an error, not silently skipped.
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "S1"
    ws.append(["Date", "Amount"])
    ws.append([dt.datetime(2026, 1, 1), None])
    ws.append([dt.datetime(2026, 12, 31), None])
    buf = io.BytesIO(); wb.save(buf); data = buf.getvalue()

    records = parse_text_lines(["Customer Name: GRP\nINV1   15/06/2026   500.00"],
                               group_label="Customer Name",
                               field_names=["Inv No", "Inv Date", "Amount"])
    cfg = _config()  # default action = skip
    rows = aggregate(records, cfg.source, aggregation_keys=["group", "date"])
    report = build_plan(excel_reader.open_workbook(data), rows, cfg)
    item = report.items[0]
    assert item.status == Status.ERROR
    assert any("WITHIN the Excel date range" in m for m in item.messages)
    assert report.has_blocking_errors


def _wb_two_branches(year=2026):
    """Two branch sheets BN and BD, each DATE/CHALLAN columns, real dates."""
    wb = openpyxl.Workbook()
    bn = wb.active; bn.title = "BN"
    bd = wb.create_sheet("BD")
    other = wb.create_sheet("Consolidated")
    for ws in (bn, bd):
        ws.append(["DATE", "CHALLAN"])
        for d in (dt.date(year, 5, 15), dt.date(year, 5, 16)):
            c = ws.cell(row=ws.max_row + 1, column=1, value=dt.datetime(d.year, d.month, d.day))
            c.number_format = "d-mmm"
            ws.cell(row=c.row, column=2, value=None)
    other.append(["DATE", "BN", "BD"])
    other.append([dt.datetime(year, 5, 15), None, None])
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def _records_two():
    text = """\
Customer Name: BARANAGAR
A1   15/05/2026   100.00
Customer Name: BEADON STREET
B1   15/05/2026   200.00
"""
    return parse_text_lines([text], group_label="Customer Name",
                            field_names=["Inv No", "Inv Date", "Amount"])


def _cfg_two():
    return AppConfig(
        source=SourceMapping(date_field="Inv Date", amount_field="Amount",
                             sum_duplicate_dates=True, date_interpretation="dmy"),
        group_to_sheet={"BARANAGAR": "BN", "BEADON STREET": "BD"},
        sheets={
            "BN": SheetMapping(sheet_name="BN", header_row=1,
                               date_target_mode=TargetMode.COLUMN_LETTER, date_column="A",
                               amount_target_mode=TargetMode.COLUMN_LETTER, amount_column="B"),
            "BD": SheetMapping(sheet_name="BD", header_row=1,
                               date_target_mode=TargetMode.COLUMN_LETTER, date_column="A",
                               amount_target_mode=TargetMode.COLUMN_LETTER, amount_column="B"),
        },
        write_rules=WriteRules(write_action=WriteAction.REPLACE, aggregation_keys=["group", "date"]),
    )


def test_sheet_per_group_writes_only_to_assigned_sheet():
    from src.excel_writer import apply_writes
    from src.validator import plan_to_write_ops
    data = _wb_two_branches(2026)
    cfg = _cfg_two()
    rows = aggregate(_records_two(), cfg.source, aggregation_keys=["group", "date"])
    report = build_plan(excel_reader.open_workbook(data), rows, cfg)
    bn_item = [i for i in report.items if i.group == "BARANAGAR"][0]
    bd_item = [i for i in report.items if i.group == "BEADON STREET"][0]
    assert bn_item.sheet == "BN" and bn_item.status == Status.READY
    assert bd_item.sheet == "BD" and bd_item.status == Status.READY

    out, _ = apply_writes(data, plan_to_write_ops(report, cfg))
    ow = openpyxl.load_workbook(io.BytesIO(out))
    # Each customer's value landed only in its own sheet at the matched row.
    assert ow["BN"].cell(row=bn_item.matched_row, column=2).value == 100.0
    assert ow["BD"].cell(row=bd_item.matched_row, column=2).value == 200.0
    # Consolidated must be untouched.
    assert ow["Consolidated"]["B2"].value is None
    assert ow["Consolidated"]["C2"].value is None


def test_date_diagnostics_ok_per_sheet():
    from src.validator import date_diagnostics
    data = _wb_two_branches(2026)
    cfg = _cfg_two()
    rows = aggregate(_records_two(), cfg.source, aggregation_keys=["group", "date"])
    diags = {d["worksheet"]: d for d in date_diagnostics(excel_reader.open_workbook(data), cfg, rows)}
    assert diags["BN"]["status"] == "OK"
    assert diags["BN"]["customers"] == ["BARANAGAR"]
    assert diags["BN"]["excel_min"] == dt.date(2026, 5, 15)
    assert diags["BD"]["customers"] == ["BEADON STREET"]
    assert diags["BN"]["matched"] == 1 and diags["BN"]["missing"] == 0


def test_date_diagnostics_no_range_message_for_old_workbook():
    from src.validator import date_diagnostics, NO_RANGE_MESSAGE
    data = _wb_two_branches(2021)   # workbook has 2021 dates, PDF is 2026
    cfg = _cfg_two()
    rows = aggregate(_records_two(), cfg.source, aggregation_keys=["group", "date"])
    diags = {d["worksheet"]: d for d in date_diagnostics(excel_reader.open_workbook(data), cfg, rows)}
    assert diags["BN"]["status"] == NO_RANGE_MESSAGE
    assert diags["BD"]["status"] == NO_RANGE_MESSAGE


def test_excel_date_debug_export_shows_real_values():
    data, row = _wb_with_format("d-mmm")
    cfg = _config()
    debug = excel_date_debug_rows(excel_reader.open_workbook(data), cfg)
    target = [d for d in debug if d["row"] == row][0]
    assert target["worksheet"] == "S1"
    assert target["cell"] == f"A{row}"
    assert target["number_format"] == "d-mmm"
    assert target["cell_data_type"] == "d"        # real date, not string
    assert target["parsed_date"] == "2026-01-04"  # real year, not display
    assert target["parse_status"] == "ok"
