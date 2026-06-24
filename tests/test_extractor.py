"""Tests for extraction-related parsing of grouped rows from sample text.

These tests use plain text (no external PDF/OCR dependency) so they run
anywhere, exercising the generic, non-hardcoded parsing logic.
"""
from src import parser


# Note: the customer/group names, labels and columns used below are arbitrary
# sample values supplied at test time. They are NOT referenced anywhere in the
# application code.

SAMPLE_PAGE_1 = """\
Report Title
Customer Name: ALPHA STORE
Inv No    Inv Date    Total Amount    Returns
A1        15/05/2026  630.00          0.00
A2        15/05/2026  12,705.00       100.00
Customer Name: BETA TRADERS
B1        16/05/2026  1,000.00        0.00
"""

SAMPLE_PAGE_2 = """\
Inv No    Inv Date    Total Amount    Returns
B2        17/05/2026  2,500.00        50.00
Customer Name: GAMMA LTD
G1        18/05/2026  9,999.00        0.00
"""


def test_extract_grouped_rows_from_text():
    records = parser.parse_text_lines(
        [SAMPLE_PAGE_1],
        group_label="Customer Name",
        field_names=["Inv No", "Inv Date", "Total Amount", "Returns"],
    )
    groups = parser.list_detected_groups(records)
    assert "ALPHA STORE" in groups
    assert "BETA TRADERS" in groups
    # Two data rows belong to ALPHA STORE.
    alpha_rows = [r for r in records if r.group == "ALPHA STORE" and not r.ignored]
    assert len(alpha_rows) == 2


def test_group_continues_across_page_break():
    records = parser.parse_text_lines(
        [SAMPLE_PAGE_1, SAMPLE_PAGE_2],
        group_label="Customer Name",
        field_names=["Inv No", "Inv Date", "Total Amount", "Returns"],
    )
    # BETA TRADERS started on page 1 and continues onto page 2 (row B2) before a
    # new Customer Name (GAMMA LTD) appears.
    beta_rows = [r for r in records if r.group == "BETA TRADERS" and not r.ignored]
    pages = sorted({r.page for r in beta_rows})
    assert pages == [1, 2]
    assert any(r.page == 2 for r in beta_rows)


def test_repeated_header_is_ignored():
    records = parser.parse_text_lines(
        [SAMPLE_PAGE_1, SAMPLE_PAGE_2],
        group_label="Customer Name",
        field_names=["Inv No", "Inv Date", "Total Amount", "Returns"],
    )
    # The header line "Inv No Inv Date ..." appears on both pages; it must never
    # be turned into a data record with a real amount.
    for r in records:
        assert r.fields.get("Inv No") != "Inv No"


def test_auto_group_label_detection():
    cands = parser.detect_group_label_candidates([SAMPLE_PAGE_1])
    assert any("Customer Name" == c for c in cands)


def test_smart_parse_switches_when_tables_lose_groups():
    # Simulate a PDF whose group labels sit OUTSIDE the tables: the detected
    # table has clean rows but no "Customer Name: ..." marker, while the page
    # text does contain the markers.
    tables = [{
        "page": 1,
        "rows": [
            ["Inv No", "Inv Date", "Total Amount", "Returns"],
            ["A1", "15/05/2026", "630.00", "0.00"],
            ["A2", "15/05/2026", "12,705.00", "0.00"],
        ],
        "ocr_confidence": None,
    }]
    # Table mode alone finds rows but zero groups.
    table_only = parser.parse_tables(tables, group_label="Customer Name")
    assert len(parser.list_detected_groups(table_only)) == 0

    records, mode_used, switched = parser.smart_parse(
        tables, [SAMPLE_PAGE_1], group_label="Customer Name", preferred="tables"
    )
    assert switched is True
    assert mode_used == "text"
    assert "ALPHA STORE" in parser.list_detected_groups(records)


def test_smart_parse_keeps_tables_when_groups_present():
    # When the table rows themselves carry the group marker, no switch happens.
    tables = [{
        "page": 1,
        "rows": [
            ["Customer Name: ALPHA STORE"],
            ["Inv No", "Inv Date", "Amount"],
            ["A1", "15/05/2026", "630.00"],
        ],
        "ocr_confidence": None,
    }]
    records, mode_used, switched = parser.smart_parse(
        tables, ["irrelevant text"], group_label="Customer Name", preferred="tables"
    )
    assert switched is False
    assert mode_used == "tables"
    assert "ALPHA STORE" in parser.list_detected_groups(records)
