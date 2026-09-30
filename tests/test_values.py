import datetime as dt

import pytest

from src import values

D = dt.date(2026, 8, 28)


@pytest.mark.parametrize("text,expected", [
    ("8,405.00", 8405.0), ("12,34,567.00", 1234567.0), ("₹ 1,23,456.78", 123456.78),
    ("(500.00)", -500.0), ("500.00-", -500.0), ("4321", 4321.0), (1500, 1500.0),
])
def test_amounts(text, expected):
    assert values.parse_amount(text) == expected


@pytest.mark.parametrize("text", ["", "ABC-I12345", "12.5.2", "Total", "1,2,3,,4", "15 Page"])
def test_non_amounts(text):
    assert values.parse_amount(text) is None


@pytest.mark.parametrize("text", [
    "28/08/2026", "28-08-2026", "28.08.2026", "28/08/26", "28-08-26",
    "2026-08-28", "2026-08-28 00:00:00", "28 Aug 2026", "28-Aug-2026", "28 August 2026",
])
def test_date_formats_normalise(text):
    assert values.parse_date(text) == D


def test_real_dates_and_serials():
    assert values.parse_date(dt.datetime(2026, 8, 28, 13, 5)) == D
    serial = (D - dt.date(1899, 12, 30)).days
    assert values.parse_date(serial, allow_serial=True) == D
    assert values.parse_date(serial) is None          # a number is not a date unless it is an Excel cell
    assert values.parse_date(2026, allow_serial=True) is None


def test_day_first_default_and_us_option():
    assert values.parse_date("05/06/2026") == dt.date(2026, 6, 5)
    assert values.parse_date("05/06/2026", values.MDY) == dt.date(2026, 5, 6)
    assert values.detect_date_order(["05/06/2026", "28/08/2026"]) == values.DMY
    assert values.detect_date_order(["05/06/2026", "08/28/2026"]) == values.MDY


@pytest.mark.parametrize("text", ["ABC-I12345", "INV-12-05-2026-3", "28082026", "Date:", "wise", "31/02/2026"])
def test_invoice_numbers_and_noise_are_never_dates(text):
    assert values.parse_date(text) is None


def test_keys():
    assert values.label_key("Inv No") == values.label_key("InvNo") == values.label_key("inv. no")
    assert values.name_key("  north   market. ") == "NORTH MARKET"
    assert values.compact_key("NORTH MARKET") == values.compact_key("NORTHMARKET")
    assert values.name_key("NORTH MARKET") != values.name_key("NORTHMARKET")
