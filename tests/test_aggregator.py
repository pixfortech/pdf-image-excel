"""Tests for grouping/summing of duplicate group+date rows."""
import datetime as dt

from src import parser
from src.aggregator import aggregate
from src.mapping import SourceMapping


def _records():
    text = """\
Customer Name: SOME GROUP
A1   15/05/2026   630.00      0.00
A2   15/05/2026   12,705.00   100.00
A3   16/05/2026   1,000.00    0.00
"""
    return parser.parse_text_lines(
        [text], group_label="Customer Name",
        field_names=["Inv No", "Inv Date", "Amount", "Returns"],
    )


def test_sum_duplicate_group_date():
    src = SourceMapping(
        group_field="", date_field="Inv Date", amount_field="Amount",
        return_field="Returns", sum_duplicate_dates=True,
    )
    rows = aggregate(_records(), src, aggregation_keys=["group", "date"])
    # Two distinct dates -> two aggregated rows.
    assert len(rows) == 2
    same_day = [r for r in rows if r.date == dt.date(2026, 5, 15)][0]
    # 630.00 + 12,705.00 = 13,335.00 (value not hardcoded in app logic).
    assert round(same_day.amount, 2) == 13335.00
    assert round(same_day.return_amount, 2) == 100.00
    assert len(same_day.source_row_indexes) == 2


def test_no_sum_when_disabled():
    src = SourceMapping(
        date_field="Inv Date", amount_field="Amount",
        sum_duplicate_dates=False,
    )
    rows = aggregate(_records(), src, aggregation_keys=["group", "date"])
    # No summing -> each source row stays separate (3 rows).
    assert len(rows) == 3


def test_date_only_aggregation_key():
    src = SourceMapping(date_field="Inv Date", amount_field="Amount", sum_duplicate_dates=True)
    rows = aggregate(_records(), src, aggregation_keys=["date"])
    assert len(rows) == 2
