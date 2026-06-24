"""Audit report generation.

Records, for every value written to Excel, the full provenance described in the
task: source file/page/group, source rows, extracted and aggregated amounts,
target worksheet/row/column/cell, existing and new values, the write action, a
timestamp and any warnings.  Output is available as CSV and JSON.
"""
from __future__ import annotations

import csv
import datetime as _dt
import io
import json
from typing import List, Optional, Sequence

from .excel_writer import WriteOutcome
from .validator import PlanItem


AUDIT_FIELDS = [
    "timestamp",
    "source_file",
    "source_page",
    "source_group",
    "source_rows",
    "extracted_date",
    "invoice_breakup",
    "aggregated_amount",
    "excel_formula_breakup",
    "return_amount",
    "target_sheet",
    "target_row",
    "target_column",
    "target_cell",
    "existing_value",
    "new_value",
    "write_action",
    "warnings",
]


def build_audit_records(
    outcomes: Sequence[WriteOutcome],
    items: Sequence[PlanItem],
    *,
    source_file: str = "",
    timestamp: Optional[str] = None,
) -> List[dict]:
    """Pair write outcomes with their plan items to produce audit rows."""
    ts = timestamp or _dt.datetime.now().isoformat(timespec="seconds")
    # Index plan items by (sheet,row,column) for matching.
    by_target = {}
    for it in items:
        if it.target_column and it.matched_row:
            by_target[(it.sheet, it.matched_row, it.target_column)] = it

    records: List[dict] = []
    for oc in outcomes:
        it = by_target.get((oc.op.sheet_name, oc.op.row, oc.op.column))
        pages = sorted({str(r.get("_page", "")) for r in (it.source_records if it else [])})
        source_rows = [r.get("_source_text", "") for r in (it.source_records if it else [])]
        records.append({
            "timestamp": ts,
            "source_file": source_file,
            "source_page": ",".join(p for p in pages if p),
            "source_group": it.group if it else oc.op.label,
            "source_rows": " | ".join(s for s in source_rows if s),
            "extracted_date": it.date_raw if it else "",
            "invoice_breakup": it.invoice_breakup if it else "",
            "aggregated_amount": it.aggregated_amount if it else "",
            "excel_formula_breakup": it.excel_formula_breakup if it else "",
            "return_amount": it.return_amount if it else "",
            "target_sheet": oc.op.sheet_name,
            "target_row": oc.op.row,
            "target_column": oc.op.column,
            "target_cell": oc.coordinate,
            "existing_value": oc.existing_value,
            "new_value": oc.written_value if not oc.skipped else "(skipped)",
            "write_action": oc.performed_action,
            "warnings": "; ".join(it.messages) if it else oc.message,
        })
    return records


def audit_to_csv(records: Sequence[dict]) -> str:
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=AUDIT_FIELDS, extrasaction="ignore")
    writer.writeheader()
    for r in records:
        writer.writerow(r)
    return buf.getvalue()


def audit_to_json(records: Sequence[dict], indent: int = 2) -> str:
    return json.dumps(list(records), indent=indent, default=str, ensure_ascii=False)


def records_to_csv(rows: Sequence[dict], fieldnames: Optional[Sequence[str]] = None) -> str:
    """Generic dict-rows-to-CSV used for raw and aggregated exports."""
    rows = list(rows)
    if not rows:
        return ""
    if fieldnames is None:
        fieldnames = list(rows[0].keys())
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=list(fieldnames), extrasaction="ignore")
    writer.writeheader()
    for r in rows:
        writer.writerow(r)
    return buf.getvalue()
