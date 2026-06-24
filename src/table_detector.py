"""Table detection helpers.

Normalises tables coming from different sources (pdfplumber, OCR layout, or a
plain list of lists) into the common ``{"page", "rows", "ocr_confidence"}``
shape consumed by :mod:`src.parser`.  It also offers a simple whitespace-based
table reconstruction used when no structural table is available.
"""
from __future__ import annotations

import re
from typing import List, Optional, Sequence


def normalise_table(rows: Sequence[Sequence], page: int, ocr_confidence: Optional[float] = None) -> dict:
    clean_rows: List[List[str]] = []
    for row in rows:
        clean_rows.append(["" if c is None else str(c).strip() for c in row])
    return {"page": page, "rows": clean_rows, "ocr_confidence": ocr_confidence}


def drop_empty_rows(table: dict) -> dict:
    rows = [r for r in table.get("rows", []) if any(str(c).strip() for c in r)]
    return {**table, "rows": rows}


def looks_tabular(rows: Sequence[Sequence], min_cols: int = 2) -> bool:
    """True if most rows have at least ``min_cols`` populated cells."""
    if not rows:
        return False
    good = 0
    for r in rows:
        if sum(1 for c in r if str(c).strip()) >= min_cols:
            good += 1
    return good >= max(1, len(rows) // 2)


def text_to_table(text: str, page: int) -> dict:
    """Reconstruct a table from a block of text using whitespace columns.

    Lines are split on runs of two-or-more spaces, which is the typical column
    gap in fixed-width report exports.  Useful as a last-resort table builder.
    """
    rows: List[List[str]] = []
    for line in (text or "").splitlines():
        if not line.strip():
            continue
        cells = re.split(r"\s{2,}", line.rstrip())
        rows.append([c.strip() for c in cells])
    return normalise_table(rows, page)
