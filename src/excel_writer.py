"""Safe writing into existing Excel workbooks with openpyxl.

Design principles:

*   The uploaded original file is never modified in place.  Callers pass in the
    original *bytes*; this module loads a fresh workbook from those bytes, edits
    only the approved target cells, and returns new output bytes.  A separate
    backup copy of the original bytes is also produced.
*   openpyxl preserves untouched cells, formulas, styles, merged cells, column
    widths, row heights, number formats, borders and fills by virtue of loading
    and re-saving the workbook.  We only assign ``.value`` on explicitly
    approved target cells.
*   When inserting a new row we optionally copy style/number-format/formula from
    a template row so formatting is retained as far as openpyxl allows.

Nothing here references a specific sheet, column, customer or value.
"""
from __future__ import annotations

import copy
import datetime as _dt
import io
from dataclasses import dataclass, field
from typing import List, Optional

from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

from .mapping import OutputType, WriteAction


@dataclass
class WriteOp:
    """A single approved write operation."""
    sheet_name: str
    row: int
    column: int                       # 1-based
    value: object                     # numeric, str, or formula string
    action: WriteAction = WriteAction.REPLACE
    output_type: OutputType = OutputType.NUMERIC
    comment: str = ""
    insert_row: bool = False          # insert a new row at ``row`` before writing
    template_row: Optional[int] = None  # row to copy formatting from when inserting
    label: str = ""                   # human label for audit


@dataclass
class WriteOutcome:
    op: WriteOp
    coordinate: str = ""
    existing_value: object = None
    written_value: object = None
    performed_action: str = ""
    skipped: bool = False
    message: str = ""


def make_backup(original_bytes: bytes) -> bytes:
    """Return an exact copy of the original bytes to use as a backup."""
    return bytes(bytearray(original_bytes))


def _coerce_existing_numeric(value) -> Optional[float]:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    return None


def apply_writes(
    original_bytes: bytes,
    ops: List[WriteOp],
    *,
    conflict_resolver=None,
) -> (bytes, List[WriteOutcome]):
    """Apply ``ops`` to a fresh copy of the workbook loaded from ``original_bytes``.

    Returns ``(output_bytes, outcomes)``.  ``conflict_resolver`` is an optional
    callable ``(op, existing_value) -> WriteAction`` used when an op's action is
    :data:`WriteAction.ASK`.  When omitted, ASK is treated as skip-on-conflict
    for safety.
    """
    from openpyxl import load_workbook

    wb = load_workbook(io.BytesIO(bytes(original_bytes)), data_only=False)
    outcomes: List[WriteOutcome] = []

    # Group inserts by sheet and process from the bottom up so row indexes stay
    # valid while inserting.
    insert_ops = [o for o in ops if o.insert_row]
    normal_ops = [o for o in ops if not o.insert_row]

    for op in sorted(insert_ops, key=lambda o: (-o.row,)):
        _do_insert(wb, op)

    for op in normal_ops + insert_ops:
        outcomes.append(_write_one(wb, op, conflict_resolver))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), outcomes


def _do_insert(wb, op: WriteOp) -> None:
    ws = wb[op.sheet_name]
    ws.insert_rows(op.row, amount=1)
    if op.template_row is not None:
        src_row = op.template_row
        if src_row >= op.row:
            src_row += 1  # template shifted down by the insert
        for c in range(1, ws.max_column + 1):
            src = ws.cell(row=src_row, column=c)
            dst = ws.cell(row=op.row, column=c)
            if src.has_style:
                dst._style = copy.copy(src._style)
            dst.number_format = src.number_format


def _write_one(wb, op: WriteOp, conflict_resolver) -> WriteOutcome:
    ws = wb[op.sheet_name]
    cell = ws.cell(row=op.row, column=op.column)
    coordinate = f"{get_column_letter(op.column)}{op.row}"
    existing = cell.value

    outcome = WriteOutcome(op=op, coordinate=coordinate, existing_value=existing)

    action = op.action
    if action == WriteAction.ASK:
        if existing not in (None, ""):
            if conflict_resolver is not None:
                action = conflict_resolver(op, existing)
            else:
                action = WriteAction.SKIP_IF_VALUE
        else:
            action = WriteAction.REPLACE

    if action == WriteAction.SKIP_IF_VALUE and existing not in (None, ""):
        outcome.skipped = True
        outcome.performed_action = "skip"
        outcome.message = "Cell already has a value; skipped."
        return outcome

    new_value = op.value

    if action == WriteAction.ADD and op.output_type == OutputType.NUMERIC:
        base = _coerce_existing_numeric(existing) or 0.0
        try:
            new_value = base + float(op.value)
        except (TypeError, ValueError):
            new_value = op.value
            outcome.message = "Existing value not numeric; replaced instead of added."
        outcome.performed_action = "add"
    else:
        outcome.performed_action = "replace" if action != WriteAction.ADD else "add"

    cell.value = new_value
    outcome.written_value = new_value

    if op.comment:
        try:
            cell.comment = Comment(op.comment, "pdf-image-excel")
        except Exception:  # pragma: no cover
            pass

    return outcome
