"""Real-file integration check (run locally; nothing here contains business data).

    python scripts/verify_real_files.py --pdf REPORT.pdf --xlsx WORKBOOK.xlsx \
        [--profile local/profile.json] [--expect local/expect.json] [--exclude-blocked]

Runs the same pipeline as the app on YOUR files, updates the workbook, and
then independently re-checks the result with openpyxl:

* the report reconciles with every printed total (stop if not);
* each group writes only to its mapped worksheet, on the row whose DATE
  equals the invoice date, in the mapped target column, with the summed value;
* every other cell in every sheet keeps its value, formula and formatting;
* sheet list, merged cells, column widths, row heights are unchanged;
* nothing is written to a sheet that no group is mapped to;
* the untouched original is kept as a backup.

Outputs (updated workbook, audit, summary) go to ``local/out`` which is
gitignored.  ``--expect`` optionally names a JSON file of expected aggregates:
``{"aggregates": [{"group": "...", "date": "28/08/2026", "amount": 19750,
"sheet": "...", "cell": "B272"}], "grand_total": 123.0, "groups": 14}``.
"""
from __future__ import annotations

import argparse
import io
import json
import sys
from copy import copy
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src import audit, pipeline, profiles, values  # noqa: E402
from src.layout import resolve_layout  # noqa: E402
from src.profiles import ProfileStore  # noqa: E402
from src.workbook import verify_update  # noqa: E402


def _style(cell):
    return (copy(cell.font), copy(cell.fill), copy(cell.border), copy(cell.alignment),
            cell.number_format, copy(cell.protection))


def independent_checks(prepared, result) -> list:
    """Re-open both files with openpyxl and compare every cell of every sheet."""
    problems = []
    before = openpyxl.load_workbook(io.BytesIO(prepared.original))
    after = openpyxl.load_workbook(io.BytesIO(result.workbook))
    planned = {(c.sheet, c.cell): c.value for c in result.changes}

    if before.sheetnames != after.sheetnames:
        problems.append("sheet list changed")
    for ws in before.worksheets:
        new = after[ws.title]
        if sorted(map(str, ws.merged_cells.ranges)) != sorted(map(str, new.merged_cells.ranges)):
            problems.append(f"{ws.title}: merged cells changed")
        if {k: v.width for k, v in ws.column_dimensions.items()} != \
                {k: v.width for k, v in new.column_dimensions.items()}:
            problems.append(f"{ws.title}: column widths changed")
        if {k: v.height for k, v in ws.row_dimensions.items()} != \
                {k: v.height for k, v in new.row_dimensions.items()}:
            problems.append(f"{ws.title}: row heights changed")
        for row in ws.iter_rows():
            for cell in row:
                other = new[cell.coordinate]
                key = (ws.title, cell.coordinate)
                if key in planned:
                    expected = planned[key]
                    ok = other.value == expected if isinstance(expected, str) else \
                        isinstance(other.value, (int, float)) and abs(other.value - expected) < 1e-9
                    if not ok:
                        problems.append(f"{ws.title}!{cell.coordinate}: {other.value!r} != planned {expected!r}")
                elif other.value != cell.value:
                    problems.append(f"{ws.title}!{cell.coordinate}: unplanned change "
                                    f"{cell.value!r} -> {other.value!r}")
                if _style(cell) != _style(other):
                    problems.append(f"{ws.title}!{cell.coordinate}: formatting changed")

    mapped = set(prepared.profile.mapped_sheets())
    for c in result.changes:
        if c.sheet not in mapped:
            problems.append(f"write to unmapped sheet {c.sheet}")
    for row in result.rows:
        resolved = resolve_layout(prepared.wb, row.sheet, prepared.profile.layout_for(row.sheet), False)
        date_letter = resolved.date
        r = int("".join(ch for ch in row.cell if ch.isdigit()))
        cell_date = values.parse_date(after[row.sheet][f"{date_letter}{r}"].value, allow_serial=True)
        if cell_date != row.total.date:
            problems.append(f"{row.sheet}!{row.cell}: row date {cell_date} != invoice date {row.total.date}")
        match = profiles.lookup_group(prepared.profile, row.total.group)
        if match.sheet != row.sheet:
            problems.append(f"{row.total.group} written to {row.sheet}, mapped to {match.sheet}")
    return problems


def expectation_checks(prepared, expect: dict) -> list:
    problems = []
    if "groups" in expect and len(prepared.report.groups) != expect["groups"]:
        problems.append(f"expected {expect['groups']} groups, found {len(prepared.report.groups)}")
    if "grand_total" in expect:
        total = sum(i.amount or 0 for i in prepared.invoices)
        if abs(total - expect["grand_total"]) > 0.005:
            problems.append(f"grand total {total:,.2f} != expected {expect['grand_total']:,.2f}")
    rows = {(r.total.group, r.total.date, r.field): r for r in prepared.plan.rows}
    for e in expect.get("aggregates", []):
        date = values.parse_date(e["date"])
        row = rows.get((e["group"], date, e.get("field", "amount")))
        if row is None:
            problems.append(f"no plan row for {e['group']} {e['date']}")
            continue
        if abs(row.total.amount - e["amount"]) > 0.005:
            problems.append(f"{e['group']} {e['date']}: aggregate {row.total.amount} != {e['amount']}")
        for k in ("sheet", "cell"):
            if k in e and getattr(row, k) != e[k]:
                problems.append(f"{e['group']} {e['date']}: {k} {getattr(row, k)!r} != {e[k]!r}")
    return problems


def run(pdf: Path, xlsx: Path, profile_path=None, expect_path=None, exclude_blocked=False,
        out_dir=Path("local/out"), store=None) -> dict:
    store = store or ProfileStore()
    profile = profiles.from_dict(json.loads(Path(profile_path).read_text())) if profile_path else None
    prepared = pipeline.prepare(pdf.read_bytes(), pdf.name, xlsx.read_bytes(), xlsx.name, store, profile)
    summary = {
        "profile": prepared.profile.profile_name, "profile_source": prepared.profile_source,
        "groups": prepared.report.groups, "invoices": len(prepared.invoices),
        "rejected_lines": len(prepared.rejected),
        "reconciliation": audit.reconciliation_rows(prepared.reconciliation),
        "reconciled": bool(prepared.reconciliation and prepared.reconciliation.checked
                           and prepared.reconciliation.ok),
        "sheets": [vars(s) for s in prepared.plan.sheets],
        "ready": prepared.ready_count, "blocked": len(prepared.plan.blocking),
        "hard_blockers": prepared.hard_blockers, "problems": [],
    }
    expect = json.loads(Path(expect_path).read_text()) if expect_path else {}
    summary["problems"] += expectation_checks(prepared, expect)
    if prepared.hard_blockers:
        summary["problems"] += prepared.hard_blockers
        return summary
    result = pipeline.update(prepared, None, exclude_blocked=exclude_blocked)
    summary["written"] = len(result.changes)
    summary["written_sheets"] = sorted({c.sheet for c in result.changes})
    summary["problems"] += verify_update(prepared.original, result.workbook, result.changes)
    summary["problems"] += independent_checks(prepared, result)
    summary["backup_intact"] = prepared.original == xlsx.read_bytes()

    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / f"Updated_{xlsx.name}").write_bytes(result.workbook)
    (out_dir / "audit.csv").write_text(audit.to_csv(audit.audit_rows(prepared, result)))
    (out_dir / "summary.json").write_text(json.dumps(summary, indent=2, default=str))
    return summary


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--pdf", required=True, type=Path)
    ap.add_argument("--xlsx", required=True, type=Path)
    ap.add_argument("--profile", help="profile JSON (default: recognise from the local store)")
    ap.add_argument("--expect", help="JSON of expected aggregates (keep it in local/)")
    ap.add_argument("--exclude-blocked", action="store_true", help="write Ready cells even if others are blocked")
    ap.add_argument("--out", type=Path, default=Path("local/out"))
    a = ap.parse_args()
    s = run(a.pdf, a.xlsx, a.profile, a.expect, a.exclude_blocked, a.out)
    print(f"Profile: {s['profile']} ({s['profile_source']})")
    print(f"Groups: {len(s['groups'])}  Invoices: {s['invoices']}  Rejected lines: {s['rejected_lines']}")
    for line in s["reconciliation"]:
        print(f"  {line['scope']:<34} extracted {line['extracted']:>15,.2f}  printed "
              f"{line['printed'] if line['printed'] is None else format(line['printed'], ',.2f'):>15}  {line['ok']}")
    for sh in s["sheets"]:
        print(f"  {sh['group']:<16} -> {sh['sheet']:<5} {sh['dates_matched']}/{sh['dates_total']:<3} "
              f"{sh['target']:<14} {sh['status']} {sh['note']}")
    print(f"Ready {s['ready']}  Blocked {s['blocked']}  Written {s.get('written', 0)} "
          f"in {', '.join(s.get('written_sheets', []))}")
    print("PASS" if not s["problems"] and s["reconciled"] else "FAIL")
    for p in s["problems"][:50]:
        print("  -", p)
    sys.exit(0 if not s["problems"] and s["reconciled"] else 1)


if __name__ == "__main__":
    main()
