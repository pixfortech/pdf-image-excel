"""Tests for amount/date parsing and total-line detection."""
import datetime as dt

from src import parser, utils


def test_parse_amount_with_commas_and_currency():
    assert utils.parse_amount("12,705.00") == 12705.00
    assert utils.parse_amount("₹ 1,23,456.78") == 123456.78
    assert utils.parse_amount("(500.00)") == -500.00
    assert utils.parse_amount("1 234,56", thousands_sep=" ", decimal_sep=",") == 1234.56
    assert utils.parse_amount("") is None
    assert utils.parse_amount("abc") is None


def test_parse_dates_multiple_formats():
    assert utils.parse_date("15/05/2026") == dt.date(2026, 5, 15)
    assert utils.parse_date("15-05-2026") == dt.date(2026, 5, 15)
    assert utils.parse_date("2026-05-15") == dt.date(2026, 5, 15)
    assert utils.parse_date(dt.date(2026, 5, 15)) == dt.date(2026, 5, 15)
    assert utils.parse_date(dt.datetime(2026, 5, 15, 9, 30)) == dt.date(2026, 5, 15)
    assert utils.parse_date("not a date") is None


def test_total_line_detection():
    assert parser.is_total_like("Total            13,335.00")
    assert parser.is_total_like("Grand Total 100")
    assert not parser.is_total_like("A2 15/05/2026 12,705.00")


def test_total_rows_are_marked_ignored():
    text = """\
Customer Name: SOME GROUP
X1   15/05/2026   630.00
Total            630.00
"""
    records = parser.parse_text_lines(
        [text], group_label="Customer Name",
        field_names=["Inv No", "Inv Date", "Amount"],
    )
    totals = [r for r in records if r.status == "total"]
    assert len(totals) == 1
    assert totals[0].ignored is True
