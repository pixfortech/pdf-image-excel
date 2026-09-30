"""Worksheet column patterns.

A profile stores ONE common pattern (e.g. DATE / main amount / returns
headers) plus rare per-sheet exceptions.  ``resolve_layout`` finds that
pattern on a given sheet BY HEADER NAME, so sheets with the same headers work
automatically even if a column or the header row moved; only sheets where a
needed header is missing (or repeated) become exceptions.

``detect_pattern`` proposes the pattern on first use: the template is the
mapped worksheet whose headers are most shared by the other mapped worksheets,
its DATE column is the column holding dates, and a returns target is detected
only when a header uniquely matches the report's returns field name.  The main
amount target is never guessed.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from openpyxl.utils import column_index_from_string

from . import values
from .profiles import ColumnRef, SheetLayout
from .workbook import Workbook

_HEADER_SEARCH_ROWS = 15


@dataclass
class Resolved:
    sheet: str
    header_row: int
    date: Optional[str] = None
    amount: Optional[str] = None
    returns: Optional[str] = None
    problems: List[str] = field(default_factory=list)
    headers: Dict[str, str] = field(default_factory=dict)

    @property
    def ok(self) -> bool:
        return not self.problems

    def describe(self, letter: Optional[str]) -> str:
        return f"{letter} ({self.headers.get(letter) or 'blank'})" if letter else "—"


def _find(headers: Dict[str, str], ref: ColumnRef) -> Tuple[Optional[str], str]:
    """Locate a column: by header name when saved (the letter only breaks a
    tie between repeated headers), otherwise by letter."""
    if ref.header:
        key = values.label_key(ref.header)
        hits = [l for l, text in headers.items() if text and values.label_key(text) == key]
        if len(hits) == 1:
            return hits[0], ""
        if len(hits) > 1:
            if ref.letter in hits:
                return ref.letter, ""
            return None, f"'{ref.header}' appears in columns {', '.join(hits)}"
        return None, f"no '{ref.header}' column"
    if ref.letter:
        try:
            column_index_from_string(ref.letter)
            return ref.letter, ""
        except ValueError:
            return None, f"invalid column '{ref.letter}'"
    return None, "not set"


def resolve_layout(wb: Workbook, sheet: str, layout: SheetLayout, need_returns: bool) -> Resolved:
    """Find the pattern's columns on ``sheet``.  Tries the saved header row,
    then (for header-named patterns) the first rows of the sheet."""
    rows = [layout.header_row] + [r for r in range(1, _HEADER_SEARCH_ROWS + 1) if r != layout.header_row]
    searchable = bool(layout.date.header and layout.amount.header)
    first: Optional[Resolved] = None
    for row in rows:
        headers = dict(wb.headers(sheet, row))
        res = Resolved(sheet, row, headers=headers)
        res.date, p_date = _find(headers, layout.date)
        res.amount, p_amount = _find(headers, layout.amount)
        p_returns = ""
        if need_returns:
            res.returns, p_returns = _find(headers, layout.returns)
        res.problems = [f"{name}: {p}" for name, p in
                        (("date", p_date), ("amount", p_amount), ("returns", p_returns)) if p]
        if res.ok:
            return res
        first = first or res
        if not searchable:
            break
    return first


def _returns_match(header: str, returns_label: str) -> bool:
    a, b = values.label_key(header), values.label_key(returns_label)
    return len(a) >= 4 and len(b) >= 4 and (a.startswith(b) or b.startswith(a))


@dataclass
class PatternSuggestion:
    template_sheet: str
    layout: SheetLayout


def detect_pattern(wb: Workbook, sheets: Sequence[str], returns_label: str = "") -> Optional[PatternSuggestion]:
    """Propose the common pattern from the mapped worksheets (see module doc)."""
    found = {}
    for sheet in dict.fromkeys(sheets):
        if not wb.has_sheet(sheet):
            continue
        guess = wb.suggest_layout(sheet)
        if guess.date.is_set:
            keys = {values.label_key(t) for _, t in wb.headers(sheet, guess.header_row) if t}
            found[sheet] = (guess, keys)
    if not found:
        return None
    template = max(found, key=lambda s: sum(len(found[s][1] & found[o][1]) for o in found if o != s))
    layout = found[template][0]
    if returns_label:
        headers = dict(wb.headers(template, layout.header_row))
        hits = [l for l, t in headers.items() if t and _returns_match(t, returns_label)
                and not wb.is_formula_column(template, l, layout.header_row)]
        if len(hits) == 1:
            layout.returns = ColumnRef(hits[0], headers[hits[0]])
    return PatternSuggestion(template, layout)
