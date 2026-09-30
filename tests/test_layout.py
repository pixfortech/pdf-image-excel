"""The common worksheet pattern: resolved by header name on every sheet,
detected automatically from the mapped branch sheets on first use."""
import io

import openpyxl

from src.layout import detect_pattern, resolve_layout
from src.profiles import ColumnRef, SheetLayout
from src.workbook import Workbook
from tests.conftest import BRANCH_LAYOUT, make_workbook


def _wb(**edits):
    data = make_workbook(["S1", "S2", "S3"], other_layout=["K1"])
    if edits:
        book = openpyxl.load_workbook(io.BytesIO(data))
        for ref, value in edits.items():
            sheet, cell = ref.split("__")
            book[sheet][cell] = value
        buf = io.BytesIO(); book.save(buf); data = buf.getvalue()
    return Workbook(data)


def test_pattern_resolves_by_header_on_every_compatible_sheet():
    wb = _wb()
    for sheet in ("S1", "S2", "S3"):
        r = resolve_layout(wb, sheet, BRANCH_LAYOUT, need_returns=True)
        assert (r.ok, r.date, r.amount, r.returns) == (True, "A", "B", "C")


def test_moved_column_still_resolves_by_header():
    wb = _wb(S2__B1="RETURN", S2__C1="CHALLAN")             # columns swapped on one sheet
    r = resolve_layout(wb, "S2", BRANCH_LAYOUT, need_returns=True)
    assert (r.amount, r.returns) == ("C", "B")


def test_different_sheet_is_an_exception_with_a_clear_reason():
    r = resolve_layout(_wb(), "K1", BRANCH_LAYOUT, need_returns=False)
    assert not r.ok and r.problems == ["amount: no 'CHALLAN' column"]


def test_returns_header_is_only_required_when_returns_will_be_written():
    wb = _wb(S1__C1="NOTES")
    assert resolve_layout(wb, "S1", BRANCH_LAYOUT, need_returns=False).ok
    assert not resolve_layout(wb, "S1", BRANCH_LAYOUT, need_returns=True).ok


def test_repeated_header_uses_saved_letter_or_asks():
    wb = _wb(S1__D1="CHALLAN")                              # CHALLAN twice (B and D)
    assert resolve_layout(wb, "S1", BRANCH_LAYOUT, need_returns=False).amount == "B"
    by_name_only = SheetLayout(1, ColumnRef(header="DATE"), ColumnRef(header="CHALLAN"))
    r = resolve_layout(wb, "S1", by_name_only, need_returns=False)
    assert not r.ok and "appears in columns B, D" in r.problems[0]


def test_detects_pattern_from_mapped_branch_sheets_not_other_sheets():
    wb = _wb()
    # Consolidated is not mapped, K1 is; the template must be a branch sheet.
    s = detect_pattern(wb, ["K1", "S1", "S2", "S3"], returns_label="Returns")
    assert s.template_sheet == "S1"
    assert (s.layout.header_row, s.layout.date.letter, s.layout.date.header) == (1, "A", "DATE")
    assert (s.layout.returns.letter, s.layout.returns.header) == ("C", "RETURN")
    assert not s.layout.amount.is_set                        # the main amount target is never guessed
