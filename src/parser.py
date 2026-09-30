"""Layout-aware parsing of grouped reports into semantic invoice rows.

The report shape this handles (and which a saved profile describes) is:

    <preamble: title, "From Date: .. To Date: ..">
    <Group Label>: <group name>              e.g. "Customer Name: NORTH MARKET"
    <column header line>                     e.g. "InvNo InvDate Total Amount Returns"
    <data rows>                              continue across pages until the next group
    <group total label> : <amount(s)>        e.g. "Customer Total : 1,23,456.00"
    ...
    <grand total label> : <amount>           e.g. "Total : 9,87,654.00"

Columns are identified by the header's POSITION on the page (ruled header
cells when present, otherwise word gaps), and data words are assigned to the
column they sit under.  Fields therefore carry their printed header label
("Inv Date", "Total Amount") — never a token index — so a blank Returns cell
cannot shift an amount into the wrong field, and an invoice number cannot land
in the date field.  Printed totals are not data: they are kept to reconcile
the extracted invoices.  Nothing here knows any customer, branch or sheet name.
"""
from __future__ import annotations

import datetime as _dt
import math
import re
import statistics
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from . import values
from .extractor import Document, Page, Word

_TOLERANCE = 0.005  # money reconciliation tolerance
_MIN_TYPE_SHARE = 0.9


# ---------------------------------------------------------------------------
# Result types
# ---------------------------------------------------------------------------

@dataclass
class Line:
    page: int
    words: List[Word]

    @property
    def text(self) -> str:
        return " ".join(w.text for w in self.words)

    @property
    def yc(self) -> float:
        return sum(w.yc for w in self.words) / len(self.words)


@dataclass
class Column:
    label: str
    lo: float
    hi: float

    @property
    def key(self) -> str:
        return values.label_key(self.label)


@dataclass
class SourceRow:
    """A line inside a group, split into columns (not yet validated)."""
    group: str
    page: int
    values: Dict[str, str]
    text: str


@dataclass
class Summary:
    """A printed total line: label followed only by amounts."""
    label: str
    amounts: List[float]
    page: int
    text: str
    group: Optional[str]


@dataclass
class IgnoredLine:
    page: int
    text: str
    reason: str


@dataclass
class ColumnStats:
    label: str
    non_empty: int = 0
    dates: int = 0
    amounts: int = 0

    @property
    def kind(self) -> str:
        if self.non_empty == 0:
            return "empty"
        if self.dates / self.non_empty >= _MIN_TYPE_SHARE:
            return "date"
        if self.amounts / self.non_empty >= _MIN_TYPE_SHARE:
            return "amount"
        return "text"


@dataclass
class ParsedReport:
    group_label: str = ""
    title: str = ""
    columns: List[str] = field(default_factory=list)
    groups: List[str] = field(default_factory=list)
    rows: List[SourceRow] = field(default_factory=list)
    group_total_label: str = ""
    group_totals: Dict[str, List[Summary]] = field(default_factory=dict)
    grand_total: Optional[Summary] = None
    period: Optional[Tuple[_dt.date, Optional[_dt.date]]] = None
    date_order: str = values.DMY
    column_stats: Dict[str, ColumnStats] = field(default_factory=dict)
    ignored: List[IgnoredLine] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    def column_kind(self, label: str) -> str:
        key = values.label_key(label)
        for name, stats in self.column_stats.items():
            if values.label_key(name) == key:
                return stats.kind
        return "missing"

    def has_column(self, label: str) -> bool:
        key = values.label_key(label)
        return any(values.label_key(c) == key for c in self.columns)


@dataclass
class Invoice:
    group: str
    date: _dt.date
    amount: Optional[float]
    returns: Optional[float]
    reference: str
    page: int
    text: str


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------

def parse(
    doc: Document,
    *,
    group_label: str = "",
    header_labels: Sequence[str] = (),
    group_total_label: str = "",
    date_order: str = values.DMY,
) -> ParsedReport:
    """Parse ``doc``.  Labels come from a saved profile when known; otherwise
    they are detected from the document's structure."""
    report = ParsedReport()
    pages = [(page, _lines(page)) for page in doc.pages]
    all_lines = [line for _, lines in pages for line in lines]

    header_key = values.label_key(" ".join(header_labels)) if header_labels else ""
    if not header_key or not any(values.label_key(l.text) == header_key for l in all_lines):
        if header_key:
            report.warnings.append(
                "The saved column header was not found; the header was re-detected from this document.")
        header_key = _detect_header_key(all_lines)
    report.group_label = group_label or _detect_group_label(all_lines, header_key)
    group_key = values.label_key(report.group_label)

    group: Optional[str] = None
    columns: Optional[List[Column]] = None
    summaries: List[Summary] = []
    preamble: List[Line] = []

    for page, lines in pages:
        for line in lines:
            if header_key and values.label_key(line.text) == header_key:
                columns = _columns_for(line, page, header_labels)
                if not report.columns:
                    report.columns = [c.label for c in columns]
                continue
            label_value = _label_value(line.text)
            if label_value and group_key and values.label_key(label_value[0]) == group_key:
                group = label_value[1]
                if group not in report.groups:
                    report.groups.append(group)
                continue
            summary = _summary(line, group)
            if summary:
                summaries.append(summary)
                continue
            if group is None:
                preamble.append(line)
                report.ignored.append(IgnoredLine(line.page, line.text, "before the first group"))
                continue
            if columns is None:
                report.ignored.append(IgnoredLine(line.page, line.text, "no column header yet"))
                continue
            report.rows.append(SourceRow(group, line.page, _assign(line, columns), line.text))

    if not report.groups:
        report.warnings.append("No groups were detected (group label not found).")
    if not report.columns:
        report.warnings.append("No column header line was detected.")

    report.period = _period(preamble)
    report.title = next((l.text for l in preamble
                         if not re.search(r"[\d:]", l.text) and len(l.words) >= 2), "")
    report.date_order = _resolve_order(date_order, report.rows)
    report.column_stats = _column_stats(report)
    _classify_summaries(report, summaries, group_total_label)
    return report


def _lines(page: Page) -> List[Line]:
    """Cluster words into visual lines by vertical centre."""
    lines: List[Line] = []
    for word in sorted(page.words, key=lambda w: (w.yc, w.x0)):
        if lines and abs(word.yc - lines[-1].yc) <= 0.5 * max(word.height, lines[-1].words[0].height):
            lines[-1].words.append(word)
        else:
            lines.append(Line(page.number, [word]))
    for line in lines:
        line.words.sort(key=lambda w: w.x0)
    return lines


_LABEL_VALUE_RE = re.compile(r"^\s*([A-Za-z][A-Za-z .]*?)\s*:\s*(.*?)\s*$")


def _label_value(text: str) -> Optional[Tuple[str, str]]:
    m = _LABEL_VALUE_RE.match(text)
    return (m.group(1).strip(), m.group(2).strip()) if m else None


def _summary(line: Line, group: Optional[str]) -> Optional[Summary]:
    """``<letters label> [:] <amount> [<amount> ...]`` — a printed total."""
    words = [w.text for w in line.words]
    label_words = []
    i = 0
    while i < len(words) and re.fullmatch(r"[A-Za-z][A-Za-z.]*:?", words[i]):
        label_words.append(words[i].rstrip(":"))
        i += 1
    if i < len(words) and words[i] == ":":
        i += 1
    amounts = [values.parse_amount(w) for w in words[i:]]
    if not label_words or not amounts or any(a is None for a in amounts) or len(amounts) > 3:
        return None
    return Summary(" ".join(label_words), amounts, line.page, line.text, group)


def _is_header_like(line: Line) -> bool:
    return (len(line.words) >= 2 and ":" not in line.text
            and all(re.search(r"[A-Za-z]", w.text) and not re.search(r"\d", w.text)
                    for w in line.words))


def _detect_header_key(lines: List[Line]) -> str:
    """The header is the header-like line most often followed by a data line
    (a line containing a date that is not a ``label: value`` line)."""
    scores: Dict[str, int] = {}
    for i, line in enumerate(lines):
        if not _is_header_like(line):
            continue
        nxt = lines[i + 1] if i + 1 < len(lines) else None
        if nxt and nxt.page == line.page and not _label_value(nxt.text) \
                and any(values.is_date(w.text) for w in nxt.words):
            key = values.label_key(line.text)
            scores[key] = scores.get(key, 0) + 1
    return max(scores, key=lambda k: scores[k]) if scores else ""


def _detect_group_label(lines: List[Line], header_key: str) -> str:
    """The group label is the ``label: value`` line that most often
    introduces a column header (e.g. ``Customer Name: X`` above the header)."""
    counts: Dict[str, List[str]] = {}
    for i, line in enumerate(lines):
        lv = _label_value(line.text)
        if not lv or not lv[1] or values.is_amount(lv[1]):
            continue
        following = lines[i + 1:i + 3]
        if any(values.label_key(f.text) == header_key for f in following):
            counts.setdefault(values.label_key(lv[0]), []).append(lv[0])
    if not counts:
        return ""
    best = max(counts, key=lambda k: len(counts[k]))
    return statistics.mode(counts[best])


def _columns_for(line: Line, page: Page, header_labels: Sequence[str]) -> List[Column]:
    ruled = _ruled_columns(line, page)
    if ruled:
        return ruled
    groups = _group_by_labels(line.words, header_labels) or _group_by_gaps(line.words)
    spans = [(" ".join(w.text for w in g), g[0].x0, g[-1].x1) for g in groups]
    columns = []
    for i, (label, x0, x1) in enumerate(spans):
        lo = -math.inf if i == 0 else (spans[i - 1][2] + x0) / 2
        hi = math.inf if i == len(spans) - 1 else (x1 + spans[i + 1][1]) / 2
        columns.append(Column(label, lo, hi))
    return columns


def _ruled_columns(line: Line, page: Page) -> Optional[List[Column]]:
    """Use the boxed header cells drawn in the PDF when they line up."""
    for row in page.ruled_rows:
        top, bottom = min(b[1] for b in row), max(b[3] for b in row)
        if not (top - 2 <= line.yc <= bottom + 2):
            continue
        columns, used = [], 0
        for x0, _t, x1, _b in sorted(row, key=lambda b: b[0]):
            inside = [w for w in line.words if x0 <= w.xc <= x1]
            used += len(inside)
            if inside:
                columns.append(Column(" ".join(w.text for w in inside), x0, x1))
        if used == len(line.words) and len(columns) >= 2:
            return columns
    return None


def _group_by_labels(words: List[Word], labels: Sequence[str]) -> Optional[List[List[Word]]]:
    """Group header words so they spell the saved profile's column labels."""
    if not labels:
        return None
    groups, i = [], 0
    for label in labels:
        target, current = values.label_key(label), []
        while i < len(words) and len(values.label_key("".join(w.text for w in current))) < len(target):
            current.append(words[i])
            i += 1
        if values.label_key("".join(w.text for w in current)) != target:
            return None
        groups.append(current)
    return groups if i == len(words) else None


def _group_by_gaps(words: List[Word]) -> List[List[Word]]:
    """Words separated by roughly one space belong to the same header label."""
    groups = [[words[0]]]
    for prev, word in zip(words, words[1:]):
        if word.x0 - prev.x1 < 0.5 * max(prev.height, word.height):
            groups[-1].append(word)
        else:
            groups.append([word])
    return groups


def _assign(line: Line, columns: List[Column]) -> Dict[str, str]:
    cells: Dict[str, List[str]] = {c.label: [] for c in columns}
    for word in line.words:
        inside = [c for c in columns if c.lo <= word.xc <= c.hi]
        column = inside[0] if inside else min(
            columns, key=lambda c: min(abs(word.xc - c.lo), abs(word.xc - c.hi)))
        cells[column.label].append(word.text)
    return {label: " ".join(parts) for label, parts in cells.items()}


def _period(preamble: List[Line]) -> Optional[Tuple[_dt.date, Optional[_dt.date]]]:
    for line in preamble:
        dates = [d for d in (values.parse_date(w.text) for w in line.words) if d]
        if len(dates) == 2:
            return min(dates), max(dates)
        if len(dates) == 1:
            return dates[0], None
    return None


def _resolve_order(order: str, rows: List[SourceRow]) -> str:
    if order != values.AUTO:
        return order
    return values.detect_date_order(v for r in rows for v in r.values.values())


def _column_stats(report: ParsedReport) -> Dict[str, ColumnStats]:
    """Type each column from data-like rows (rows containing a date)."""
    stats = {label: ColumnStats(label) for label in report.columns}
    for row in report.rows:
        if not any(values.parse_date(v, report.date_order) for v in row.values.values()):
            continue
        for label, text in row.values.items():
            s = stats.setdefault(label, ColumnStats(label))
            if not text.strip():
                continue
            s.non_empty += 1
            s.dates += values.parse_date(text, report.date_order) is not None
            s.amounts += values.is_amount(text)
    return stats


def _classify_summaries(report: ParsedReport, summaries: List[Summary], group_total_label: str) -> None:
    if not summaries:
        return
    counts: Dict[str, int] = {}
    for s in summaries:
        counts[values.label_key(s.label)] = counts.get(values.label_key(s.label), 0) + 1
    if group_total_label:
        gt_key = values.label_key(group_total_label)
    else:
        first_seen = {}
        for i, s in enumerate(summaries):
            first_seen.setdefault(values.label_key(s.label), i)
        gt_key = max(counts, key=lambda k: (counts[k], -first_seen[k])) if report.groups else ""
    for s in summaries:
        if gt_key and values.label_key(s.label) == gt_key and s.group:
            report.group_totals.setdefault(s.group, []).append(s)
            report.group_total_label = report.group_total_label or s.label
    others = [s for s in summaries if values.label_key(s.label) != gt_key]
    if others:
        report.grand_total = others[-1]
        for s in others[:-1]:
            report.ignored.append(IgnoredLine(s.page, s.text, "unclassified total line"))


# ---------------------------------------------------------------------------
# Semantic validation: only genuine invoice rows enter the pipeline
# ---------------------------------------------------------------------------

def invoices(
    report: ParsedReport,
    *,
    date_field: str,
    amount_field: str,
    returns_field: str = "",
    reference_field: str = "",
) -> Tuple[List[Invoice], List[IgnoredLine]]:
    """Validate rows against the semantic fields.  A row is an invoice only if
    its date field strictly parses as a date and its amount fields are
    numeric (or blank)."""
    out: List[Invoice] = []
    rejected: List[IgnoredLine] = []
    for row in report.rows:
        get = lambda label: _value(row, label) if label else ""
        date_text, amount_text, returns_text = get(date_field), get(amount_field), get(returns_field)
        date = values.parse_date(date_text, report.date_order)
        amount = values.parse_amount(amount_text) if amount_text else None
        returns = values.parse_amount(returns_text) if returns_text else None
        reason = ""
        if date is None:
            reason = f"'{date_field}' is not a date ({date_text!r})"
        elif amount_text and amount is None:
            reason = f"'{amount_field}' is not numeric ({amount_text!r})"
        elif returns_text and returns is None:
            reason = f"'{returns_field}' is not numeric ({returns_text!r})"
        elif amount is None and returns is None:
            reason = "no amount"
        if reason:
            rejected.append(IgnoredLine(row.page, row.text, reason))
            continue
        out.append(Invoice(row.group, date, amount, returns, get(reference_field), row.page, row.text))
    return out, rejected


def _value(row: SourceRow, label: str) -> str:
    key = values.label_key(label)
    for name, text in row.values.items():
        if values.label_key(name) == key:
            return text.strip()
    return ""


# ---------------------------------------------------------------------------
# Reconciliation against printed totals
# ---------------------------------------------------------------------------

@dataclass
class ReconLine:
    scope: str
    extracted: float
    printed: Optional[float]

    @property
    def difference(self) -> Optional[float]:
        return None if self.printed is None else round(self.extracted - self.printed, 2)

    @property
    def ok(self) -> bool:
        return self.printed is not None and abs(self.extracted - self.printed) < _TOLERANCE


@dataclass
class Reconciliation:
    lines: List[ReconLine] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    @property
    def checked(self) -> bool:
        return any(l.printed is not None for l in self.lines)

    @property
    def ok(self) -> bool:
        return all(l.ok for l in self.lines if l.printed is not None)

    @property
    def failures(self) -> List[ReconLine]:
        return [l for l in self.lines if l.printed is not None and not l.ok]


def reconcile(report: ParsedReport, rows: List[Invoice]) -> Reconciliation:
    """Compare extracted invoice sums with every printed total."""
    result = Reconciliation()
    by_group: Dict[str, List[Invoice]] = {}
    for inv in rows:
        by_group.setdefault(inv.group, []).append(inv)

    printed_group_sum = 0.0
    for group in report.groups:
        invs = by_group.get(group, [])
        amount = sum(i.amount or 0 for i in invs)
        totals = report.group_totals.get(group, [])
        if not invs:
            result.warnings.append(f"Group '{group}' has no invoice rows.")
        if not totals:
            if report.group_total_label:
                result.warnings.append(f"No printed {report.group_total_label} for '{group}'.")
            result.lines.append(ReconLine(group, amount, None))
            continue
        printed = sum(t.amounts[0] for t in totals)
        printed_group_sum += printed
        result.lines.append(ReconLine(group, amount, printed))
        if any(len(t.amounts) > 1 for t in totals):
            ret = sum(i.returns or 0 for i in invs)
            result.lines.append(ReconLine(f"{group} (returns)", ret,
                                          sum(t.amounts[1] for t in totals if len(t.amounts) > 1)))

    total = sum(i.amount or 0 for i in rows)
    if report.grand_total:
        result.lines.append(ReconLine("All invoices", total, report.grand_total.amounts[0]))
        if report.group_totals:
            result.lines.append(ReconLine(f"Sum of printed {report.group_total_label}s",
                                          printed_group_sum, report.grand_total.amounts[0]))
    else:
        result.lines.append(ReconLine("All invoices", total, None))
        result.warnings.append("No grand total printed; overall total could not be reconciled.")
    return result
