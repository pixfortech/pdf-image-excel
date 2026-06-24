"""Tests for robust date parsing, normalisation and date-match warnings."""
import datetime as dt
import io

import openpyxl

from src import excel_reader, utils
from src.aggregator import aggregate
from src.mapping import (
    AppConfig, SheetMapping, SourceMapping, TargetMode, WriteAction, WriteRules,
)
from src.parser import parse_text_lines
from src.validator import Status, build_plan

TARGET = dt.date(2026, 5, 15)


def test_parse_ddmmyyyy_variants():
    for s in ("15/05/2026", "15-05-2026", "15.05.2026"):
        assert utils.parse_date(s) == TARGET


def test_parse_iso_and_datetime():
    assert utils.parse_date("2026-05-15") == TARGET
    assert utils.parse_date("2026-05-15 00:00:00") == TARGET
    assert utils.parse_date(dt.datetime(2026, 5, 15, 9, 30)) == TARGET
    assert utils.parse_date(dt.date(2026, 5, 15)) == TARGET


def test_parse_named_month():
    assert utils.parse_date("15 May 2026") == TARGET
    assert utils.parse_date("15-May-2026") == TARGET


def test_parse_two_digit_year():
    assert utils.parse_date("15/05/26") == TARGET
    assert utils.parse_date("15-05-26") == TARGET


def test_excel_serial_date():
    serial = (TARGET - dt.date(1899, 12, 30)).days  # 46157
    assert utils.parse_date(serial) == TARGET
    assert utils.excel_serial_to_date(serial) == TARGET
    # A bare small number / a year is NOT treated as a serial date.
    assert utils.parse_date(2026) is None
    assert utils.parse_date(12) is None


def test_indian_british_default_interpretation():
    # 05/06/2026 must be 5 June 2026 under the default (dmy), not 6 May.
    assert utils.parse_date("05/06/2026") == dt.date(2026, 6, 5)
    assert utils.parse_date("05/06/2026", interpretation="dmy") == dt.date(2026, 6, 5)


def test_us_interpretation():
    assert utils.parse_date("05/06/2026", interpretation="mdy") == dt.date(2026, 5, 6)
    assert utils.parse_date("05/15/2026", interpretation="mdy") == dt.date(2026, 5, 15)


def test_reject_non_date_text():
    for s in ("wise", "Date:", "sales", "From Date", "To Date", "Page No", ""):
        assert utils.parse_date(s) is None


def test_normalise_aliases_parse_date():
    assert utils.normalise_date("15 May 2026") == TARGET


# --- end-to-end: invalid rows rejected, and the >25% missing warning ---

def _records(dates):
    lines = ["Customer Name: GRP"]
    for i, d in enumerate(dates):
        lines.append(f"INV{i}   {d}   100.00")
    return parse_text_lines(["\n".join(lines)], group_label="Customer Name",
                            field_names=["Inv No", "Inv Date", "Amount"])


def _wb_with_dates(days):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "S1"
    ws.append(["Date", "Amount"])
    for d in days:
        ws.append([d, None])
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def _cfg():
    return AppConfig(
        source=SourceMapping(date_field="Inv Date", amount_field="Amount",
                             sum_duplicate_dates=True, date_interpretation="dmy"),
        group_to_sheet={"GRP": "S1"},
        sheets={"S1": SheetMapping(sheet_name="S1", header_row=1,
                                   date_target_mode=TargetMode.COLUMN_LETTER, date_column="A",
                                   amount_target_mode=TargetMode.COLUMN_LETTER, amount_column="B")},
        write_rules=WriteRules(write_action=WriteAction.REPLACE, aggregation_keys=["group", "date"]),
    )


def test_invalid_rows_do_not_enter_write_plan():
    # One good row, plus parser noise rows with non-date text.
    text = """\
Customer wise sales details
From Date: 15/05/2026 ToDate: 23/06/2026
Customer Name: GRP
INV1   15/05/2026   100.00
"""
    records = parse_text_lines([text], group_label="Customer Name",
                               field_names=["Inv No", "Inv Date", "Amount"])
    # Preamble (title, From Date) must have been skipped at parse time.
    assert all(r.group == "GRP" for r in records if not r.ignored)

    data = _wb_with_dates([dt.date(2026, 5, 15)])
    cfg = _cfg()
    rows = aggregate(records, cfg.source, aggregation_keys=["group", "date"])
    report = build_plan(excel_reader.open_workbook(data), rows, cfg)
    # Exactly one READY row; no row with a bogus date like "wise".
    ready = [i for i in report.items if i.status == Status.READY]
    assert len(ready) == 1
    assert ready[0].pdf_date_normalised == "2026-05-15"


def test_warning_when_most_dates_missing():
    # PDF dates are in 2026; the worksheet only has 2021 dates -> mostly missing.
    pdf_dates = ["15/05/2026", "16/05/2026", "17/05/2026", "18/05/2026"]
    records = _records(pdf_dates)
    data = _wb_with_dates([dt.date(2021, 1, d) for d in range(1, 6)])
    cfg = _cfg()
    rows = aggregate(records, cfg.source, aggregation_keys=["group", "date"])
    report = build_plan(excel_reader.open_workbook(data), rows, cfg)
    assert report.date_match_warnings, "expected a strong warning about missing dates"
    assert any("not found" in w for w in report.date_match_warnings)


def test_default_action_is_skip_not_insert():
    cfg = _cfg()  # default WriteRules
    assert cfg.write_rules.date_not_found_action.value == "skip"
    assert cfg.write_rules.insert_missing_date_rows is False
    records = _records(["15/05/2026"])
    data = _wb_with_dates([dt.date(2021, 1, 1)])   # date not present
    rows = aggregate(records, cfg.source, aggregation_keys=["group", "date"])
    report = build_plan(excel_reader.open_workbook(data), rows, cfg)
    # Not found and NOT inserted by default.
    assert any(i.status == Status.DATE_NOT_FOUND for i in report.items)
    assert not any(i.is_insert for i in report.items)


def test_suggest_date_columns_finds_better_column():
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "S1"
    ws.append(["Code", "Amount", "Txn Date"])     # dates live in column C, not A
    for d in (dt.date(2026, 5, 15), dt.date(2026, 5, 16)):
        ws.append(["x", 0, d])
    buf = io.BytesIO(); wb.save(buf)
    w = excel_reader.open_workbook(buf.getvalue())
    sugg = excel_reader.suggest_date_columns(w, "S1", {dt.date(2026, 5, 15), dt.date(2026, 5, 16)}, header_row=1)
    assert sugg, "expected at least one suggestion"
    assert sugg[0][0] == "C"          # column letter C is the best match
    assert sugg[0][2] == 2            # both dates matched
