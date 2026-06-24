"""Line-based and structure-aware parsing of extracted document content.

This module turns raw extracted pages (text and/or detected tables) into a flat
list of structured records.  It detects "group" sections (e.g. a line such as
``Customer Name: SOMETHING``) and assigns subsequent data rows to the most
recent group, carrying the active group across page breaks.

Crucially nothing here is tied to a particular label, customer, column, date
format or amount style.  Group labels are auto-detected as *candidates* and the
final choice is left to the user via configuration.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence

from . import utils


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class Record:
    """One detected data row plus all its provenance."""
    page: int
    group: str                       # detected group/customer/branch name
    fields: Dict[str, str] = field(default_factory=dict)  # detected field -> value
    source_text: str = ""
    ocr_confidence: Optional[float] = None
    status: str = "ok"
    warnings: List[str] = field(default_factory=list)
    ignored: bool = False
    row_index: int = -1              # stable index assigned at parse time


# A line that looks like ``<label>: <value>`` is a candidate group marker.
_GROUP_MARKER_RE = re.compile(r"^\s*([A-Za-z][A-Za-z0-9 ._/&-]{1,40}?)\s*[:\-]\s*(.+?)\s*$")

# Tokens that very commonly identify a *label* (left-hand side) introducing a
# group.  We only use these to *rank* auto-detected candidates, never to force
# a choice; the user always decides.
_GROUP_LABEL_HINTS = (
    "customer", "branch", "party", "account", "ledger", "name", "client",
    "vendor", "supplier", "dealer", "group", "shop", "outlet",
)


def looks_like_group_marker(line: str, group_label: str = "") -> Optional[tuple]:
    """Return ``(label, value)`` if *line* introduces a group, else ``None``.

    When ``group_label`` is provided, the label must match it (case-insensitive,
    ignoring surrounding spaces).  Otherwise any ``label: value`` pattern whose
    label looks like a name field is accepted.
    """
    m = _GROUP_MARKER_RE.match(line)
    if not m:
        return None
    label, value = m.group(1).strip(), m.group(2).strip()
    if not value:
        return None
    if group_label:
        if label.lower().strip() == group_label.lower().strip():
            return label, value
        return None
    # Auto mode: accept if the label resembles a known grouping label hint.
    low = label.lower()
    if any(hint in low for hint in _GROUP_LABEL_HINTS):
        return label, value
    return None


def detect_group_label_candidates(pages_text: Sequence[str]) -> List[str]:
    """Scan text and return candidate group-label strings, most frequent first."""
    counts: Dict[str, int] = {}
    for text in pages_text:
        for line in (text or "").splitlines():
            m = _GROUP_MARKER_RE.match(line)
            if not m:
                continue
            label = m.group(1).strip()
            low = label.lower()
            if any(hint in low for hint in _GROUP_LABEL_HINTS):
                counts[label] = counts.get(label, 0) + 1
    return [lbl for lbl, _ in sorted(counts.items(), key=lambda kv: -kv[1])]


# ---------------------------------------------------------------------------
# Header / footer / total detection
# ---------------------------------------------------------------------------

def _is_repeated_header(cells: Sequence[str], header_cells: Optional[Sequence[str]]) -> bool:
    if not header_cells:
        return False
    norm = [utils.normalise_whitespace(str(c)).lower() for c in cells]
    hnorm = [utils.normalise_whitespace(str(c)).lower() for c in header_cells]
    # Consider it a repeated header if the majority of non-empty cells match.
    matches = sum(1 for c in norm if c and c in hnorm)
    nonempty = sum(1 for c in norm if c)
    return nonempty > 0 and matches >= max(2, nonempty - 1)


def is_total_like(text: str, total_tokens: Optional[Sequence[str]] = None) -> bool:
    """Heuristically decide whether a line is a subtotal/total line.

    ``total_tokens`` may be supplied by configuration; when omitted a small set
    of generic words is used purely as a heuristic (still user-overridable).
    """
    tokens = total_tokens if total_tokens is not None else ("total", "subtotal", "grand total", "sum")
    low = utils.normalise_whitespace(text).lower()
    return any(low.startswith(t) or f" {t}" in f" {low}" for t in tokens) and not re.search(r"[A-Za-z]{2,}\d", low)


# ---------------------------------------------------------------------------
# Table-based parsing
# ---------------------------------------------------------------------------

def parse_tables(
    tables: Sequence[dict],
    *,
    group_label: str = "",
    total_tokens: Optional[Sequence[str]] = None,
) -> List[Record]:
    """Parse a list of detected tables into records.

    Each ``table`` is a dict with keys: ``page`` (int), ``rows`` (list of list of
    str), and optionally ``ocr_confidence``.  The first non-empty row of the
    first table that has repeated structure is treated as the header; repeated
    headers on later pages are skipped.  The active group carries across tables
    and pages until a new group marker appears.
    """
    records: List[Record] = []
    header: Optional[List[str]] = None
    active_group = ""
    idx = 0

    for table in tables:
        page = table.get("page", 0)
        rows = table.get("rows", [])
        conf = table.get("ocr_confidence")
        for raw in rows:
            cells = ["" if c is None else str(c).strip() for c in raw]
            joined = " ".join(c for c in cells if c)
            if not joined.strip():
                continue

            # A single populated cell may be a group marker line.
            marker = looks_like_group_marker(joined, group_label)
            if marker:
                active_group = marker[1]
                continue

            # Establish header from the first structured row we encounter.
            if header is None:
                header = [c if c else f"col_{i}" for i, c in enumerate(cells)]
                continue

            if _is_repeated_header(cells, header):
                continue
            if is_total_like(joined, total_tokens):
                rec = _make_record(page, active_group, header, cells, joined, conf, idx)
                rec.ignored = True
                rec.status = "total"
                rec.warnings.append("Detected as total/subtotal line")
                records.append(rec)
                idx += 1
                continue

            rec = _make_record(page, active_group, header, cells, joined, conf, idx)
            records.append(rec)
            idx += 1

    return records


def _make_record(page, group, header, cells, joined, conf, idx) -> Record:
    fields: Dict[str, str] = {}
    for i, name in enumerate(header):
        fields[name] = cells[i] if i < len(cells) else ""
    rec = Record(
        page=page,
        group=group,
        fields=fields,
        source_text=joined,
        ocr_confidence=conf,
        row_index=idx,
    )
    if not group:
        rec.warnings.append("Row has no detected group yet")
    return rec


# ---------------------------------------------------------------------------
# Free-text line parsing (fallback)
# ---------------------------------------------------------------------------

def parse_text_lines(
    pages_text: Sequence[str],
    *,
    group_label: str = "",
    row_regex: Optional[str] = None,
    field_names: Optional[Sequence[str]] = None,
    total_tokens: Optional[Sequence[str]] = None,
) -> List[Record]:
    """Parse plain text pages line by line.

    When ``row_regex`` is supplied it is used (with named or positional groups)
    to extract fields from each data line.  Otherwise rows are split on runs of
    whitespace and assigned generic ``col_N`` names.  Group markers switch the
    active group and carry across pages.
    """
    records: List[Record] = []
    active_group = ""
    idx = 0
    compiled = re.compile(row_regex) if row_regex else None

    for page_no, text in enumerate(pages_text, start=1):
        for line in (text or "").splitlines():
            if not line.strip():
                continue
            marker = looks_like_group_marker(line, group_label)
            if marker:
                active_group = marker[1]
                continue
            if is_total_like(line, total_tokens):
                rec = Record(
                    page=page_no, group=active_group,
                    fields={}, source_text=line.strip(),
                    status="total", ignored=True, row_index=idx,
                    warnings=["Detected as total/subtotal line"],
                )
                records.append(rec)
                idx += 1
                continue

            fields = _extract_line_fields(line, compiled, field_names)
            if fields is None:
                continue
            rec = Record(
                page=page_no, group=active_group,
                fields=fields, source_text=line.strip(), row_index=idx,
            )
            if not active_group:
                rec.warnings.append("Row has no detected group yet")
            records.append(rec)
            idx += 1
    return records


def _extract_line_fields(line, compiled, field_names) -> Optional[Dict[str, str]]:
    if compiled is not None:
        m = compiled.search(line)
        if not m:
            return None
        gd = m.groupdict()
        if gd:
            return {k: (v or "").strip() for k, v in gd.items()}
        return {f"col_{i}": (g or "").strip() for i, g in enumerate(m.groups())}

    # Default: split on 2+ spaces (typical of fixed-width report exports),
    # falling back to single spaces.
    parts = re.split(r"\s{2,}", line.strip())
    if len(parts) < 2:
        parts = line.strip().split()
    if len(parts) < 2:
        return None
    if field_names:
        return {
            (field_names[i] if i < len(field_names) else f"col_{i}"): p.strip()
            for i, p in enumerate(parts)
        }
    return {f"col_{i}": p.strip() for i, p in enumerate(parts)}


# ---------------------------------------------------------------------------
# Helpers for the UI
# ---------------------------------------------------------------------------

def smart_parse(
    tables: Sequence[dict],
    pages_text: Sequence[str],
    *,
    group_label: str = "",
    preferred: str = "tables",
    row_regex: Optional[str] = None,
    total_tokens: Optional[Sequence[str]] = None,
):
    """Parse with the preferred strategy, auto-switching when it loses groups.

    Returns ``(records, mode_used, switched)``.

    ``preferred`` is ``"tables"``, ``"text"`` or ``"regex"``.  When the table
    strategy yields rows but **zero** groups (a common case where group labels
    such as ``Customer Name: ...`` sit *outside* the detected tables) and parsing
    the page text *does* recover groups, this transparently falls back to
    text-line parsing and reports ``switched=True`` so the UI can warn the user.
    """
    if preferred == "regex" and row_regex:
        recs = parse_text_lines(pages_text, group_label=group_label,
                                row_regex=row_regex, total_tokens=total_tokens)
        return recs, "regex", False

    if preferred == "tables" and tables:
        recs = parse_tables(tables, group_label=group_label, total_tokens=total_tokens)
        if not list_detected_groups(recs):
            text_recs = parse_text_lines(pages_text, group_label=group_label,
                                         total_tokens=total_tokens)
            if list_detected_groups(text_recs):
                return text_recs, "text", True
        return recs, "tables", False

    recs = parse_text_lines(pages_text, group_label=group_label, total_tokens=total_tokens)
    return recs, "text", False


def list_detected_groups(records: Sequence[Record]) -> List[str]:
    """Return unique detected group names in order of first appearance."""
    seen: List[str] = []
    for r in records:
        if r.group and r.group not in seen:
            seen.append(r.group)
    return seen


def list_field_names(records: Sequence[Record]) -> List[str]:
    names: List[str] = []
    for r in records:
        for k in r.fields:
            if k not in names:
                names.append(k)
    return names
