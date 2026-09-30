"""Audit and export tables (CSV / JSON)."""
from __future__ import annotations

import csv
import datetime as _dt
import io
import json
from typing import Iterable, List

from . import values
from .plan import PlanRow


def plan_rows(rows: Iterable[PlanRow]) -> List[dict]:
    """The write preview / audit record: one row per target cell."""
    out = []
    for r in rows:
        t = r.total
        out.append({
            "customer": t.group,
            "date": values.format_date(t.date),
            "field": r.field,
            "invoice_count": len(t.invoices),
            "invoices": t.references,
            "breakup": t.breakup if r.field == "amount" else values.format_number(t.returns),
            "aggregate": t.amount if r.field == "amount" else t.returns,
            "formula": t.formula if r.field == "amount" else "",
            "worksheet": r.sheet,
            "cell": r.cell,
            "existing_value": r.existing,
            "new_value": r.new_value,
            "status": r.status,
            "note": r.note,
            "source_pages": ", ".join(sorted({str(i.page) for i in t.invoices})),
        })
    return out


def audit_rows(prepared, result) -> List[dict]:
    stamp = _dt.datetime.now().isoformat(timespec="seconds")
    rows = plan_rows(result.rows)
    for row in rows:
        row.update(timestamp=stamp, source_file=prepared.pdf_name,
                   workbook=prepared.xlsx_name, profile=prepared.profile.profile_name)
    return rows


def invoice_rows(invoices) -> List[dict]:
    return [{"customer": i.group, "date": values.format_date(i.date), "reference": i.reference,
             "amount": i.amount, "returns": i.returns, "page": i.page, "source_text": i.text}
            for i in invoices]


def reconciliation_rows(recon) -> List[dict]:
    return [{"scope": l.scope, "extracted": l.extracted, "printed": l.printed,
             "difference": l.difference, "ok": l.ok if l.printed is not None else "not printed"}
            for l in recon.lines] if recon else []


def to_csv(rows: List[dict]) -> str:
    if not rows:
        return ""
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=list(rows[0].keys()))
    writer.writeheader()
    writer.writerows(rows)
    return buf.getvalue()


def to_json(rows: List[dict]) -> str:
    return json.dumps(rows, indent=2, default=str, ensure_ascii=False)
