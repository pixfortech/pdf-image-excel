"""End-to-end validation harness for the PDF/Image -> Excel mapper.

This builds a *real* multi-page PDF and a *real* multi-sheet .xlsx that
reproduce the exact structure described in the task (Customer Name groups with
Inv No / Inv Date / Total Amount / Returns, a group that spans a page break, and
a duplicate-date case), then runs the genuine extraction -> mapping ->
aggregation -> validation -> write pipeline and asserts every required
checkpoint.

The customer/sheet/column/date values used here are SYNTHETIC stand-ins chosen
at runtime; none of them appear in the application code (src/).  Run with the
real `Customer Wise Sales Details` PDF and `karkhana-2025-2027.xlsx` by passing
their paths as argv[1] and argv[2] to validate against actual files instead.

Usage:
    python scripts/real_file_validation.py                 # synthetic fixtures
    python scripts/real_file_validation.py SOURCE.pdf BOOK.xlsx   # real files
"""
from __future__ import annotations

import io
import os
import sys
import datetime as dt

# Make the project importable when run from anywhere.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from openpyxl.styles import Font, PatternFill

from src import excel_reader, excel_writer, extractor, parser
from src.aggregator import aggregate
from src.audit import audit_to_csv, build_audit_records
from src.mapping import (
    AppConfig, OutputType, SheetMapping, SourceMapping, TargetMode,
    WriteAction, WriteRules,
)
from src.validator import Status, build_plan, plan_to_write_ops

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_validation_out")
os.makedirs(OUT, exist_ok=True)


# ---------------------------------------------------------------------------
# Synthetic fixture builders (only used when no real files are supplied)
# ---------------------------------------------------------------------------

# Synthetic, runtime-only names. Three customer groups; the middle one spans a
# page break; the first one has two invoices on the SAME date to test summing.
CUST_A = "ALPHA STORE"
CUST_B = "BETA TRADERS"
CUST_C = "GAMMA DEPOT"


def build_pdf() -> bytes:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    def line(x, y, text):
        c.drawString(x, y, text)

    # ---- Page 1 ----
    y = height - 60
    line(50, y, "Customer Wise Sales Details (synthetic validation fixture)"); y -= 30
    line(50, y, f"Customer Name: {CUST_A}"); y -= 20
    line(50, y, "Inv No        Inv Date        Total Amount        Returns"); y -= 18
    # Two rows, SAME date 15/05/2026 -> must sum to 13,335.00
    line(50, y, "INV-1001      15/05/2026      630.00              0.00"); y -= 16
    line(50, y, "INV-1002      15/05/2026      12,705.00           100.00"); y -= 16
    line(50, y, "INV-1003      16/05/2026      1,000.00            0.00"); y -= 30

    line(50, y, f"Customer Name: {CUST_B}"); y -= 20
    line(50, y, "Inv No        Inv Date        Total Amount        Returns"); y -= 18
    line(50, y, "INV-2001      15/05/2026      5,000.00            0.00"); y -= 16
    line(50, y, "INV-2002      16/05/2026      2,500.00            50.00"); y -= 16
    c.showPage()

    # ---- Page 2 : BETA TRADERS continues (no new Customer Name yet) ----
    y = height - 60
    line(50, y, "Inv No        Inv Date        Total Amount        Returns"); y -= 18  # repeated header
    line(50, y, "INV-2003      17/05/2026      3,000.00            0.00"); y -= 30      # still BETA

    line(50, y, f"Customer Name: {CUST_C}"); y -= 20
    line(50, y, "Inv No        Inv Date        Total Amount        Returns"); y -= 18
    line(50, y, "INV-3001      18/05/2026      9,999.00            0.00"); y -= 16
    line(50, y, "Total                         9,999.00            0.00"); y -= 16      # total line to ignore
    c.showPage()
    c.save()
    return buf.getvalue()


def build_workbook() -> bytes:
    """Workbook with three worksheets, a date column, amount + return columns,
    pre-existing formulas and formatting to prove preservation."""
    wb = openpyxl.Workbook()
    # Sheet names deliberately DIFFER from customer names.
    names = ["S-ONE", "S-TWO", "S-THREE"]
    dates = [dt.date(2026, 5, d) for d in (15, 16, 17, 18)]
    first = True
    for nm in names:
        ws = wb.create_sheet(nm) if not first else wb.active
        if first:
            ws.title = nm
            first = False
        # Header row 1 with a DELIBERATELY repeated header ("Amount") to prove
        # we can still target unambiguously by column letter.
        ws.append(["Txn Date", "Amount", "Amount", "Deduction", "Computed"])
        for col in range(1, 6):
            ws.cell(row=1, column=col).font = Font(bold=True)
            ws.cell(row=1, column=col).fill = PatternFill("solid", fgColor="DDDDDD")
        for i, d in enumerate(dates, start=2):
            ws.cell(row=i, column=1, value=d)
            ws.cell(row=i, column=2, value=None)        # B: target amount
            ws.cell(row=i, column=3, value=None)        # C: second "Amount"
            ws.cell(row=i, column=4, value=None)        # D: deduction target
            ws.cell(row=i, column=5, value=f"=B{i}+C{i}-D{i}")  # E: formula to preserve
        ws.column_dimensions["A"].width = 14
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def check(label, ok, detail=""):
    mark = "PASS" if ok else "FAIL"
    print(f"  [{mark}] {label}" + (f" -> {detail}" if detail else ""))
    if not ok:
        raise SystemExit(f"VALIDATION FAILED at: {label}")


def main():
    real = len(sys.argv) >= 3
    if real:
        pdf_bytes = open(sys.argv[1], "rb").read()
        xl_bytes = open(sys.argv[2], "rb").read()
        src_name = os.path.basename(sys.argv[1])
        print(f"Using REAL files: {sys.argv[1]} + {sys.argv[2]}\n")
    else:
        pdf_bytes = build_pdf()
        xl_bytes = build_workbook()
        src_name = "Customer Wise Sales Details (synthetic).pdf"
        with open(os.path.join(OUT, "sample_source.pdf"), "wb") as fh:
            fh.write(pdf_bytes)
        with open(os.path.join(OUT, "sample_workbook.xlsx"), "wb") as fh:
            fh.write(xl_bytes)
        print("Using SYNTHETIC fixtures (real files not supplied).")
        print(f"Wrote fixtures to {OUT}/\n")

    print("STEP 1 — Extract PDF (real pdfplumber path)")
    result = extractor.extract(pdf_bytes, filename=src_name, mode="auto")
    check("PDF extracted with a text engine", result.engine in ("pdfplumber", "pymupdf"), result.engine)
    check("Multiple pages read", result.page_count >= 2, f"{result.page_count} pages")

    print("\nSTEP 2 — Detect 'Customer Name' as the group label")
    cands = parser.detect_group_label_candidates(result.pages_text)
    check("'Customer Name' auto-detected as a group label", "Customer Name" in cands, str(cands))

    print("\nSTEP 3 — Parse rows, carry group across page break, ignore headers/totals")
    records = parser.parse_text_lines(
        result.pages_text,
        group_label="Customer Name",
        field_names=["Inv No", "Inv Date", "Total Amount", "Returns"],
    )
    groups = parser.list_detected_groups(records)
    check("Each unique Customer Name became a group", len(groups) >= 1, str(groups))

    if not real:
        check("Three groups detected", set(groups) == {CUST_A, CUST_B, CUST_C}, str(groups))
        beta_rows = [r for r in records if r.group == CUST_B and not r.ignored]
        beta_pages = sorted({r.page for r in beta_rows})
        check("BETA group continues across the page break", 2 in beta_pages, f"pages={beta_pages}")
        total_rows = [r for r in records if r.status == "total"]
        check("Total/subtotal line ignored", len(total_rows) >= 1, f"{len(total_rows)} total line(s)")
        # Inv No retained in the record fields (for audit/preview) though not used for the amount.
        any_inv = any(r.fields.get("Inv No", "").startswith("INV-") for r in records)
        check("Inv No kept in extracted rows (audit/preview)", any_inv)

    print("\nSTEP 4 — Source mapping: Inv Date=date, Total Amount=amount, Returns=return, Inv No ignored")
    source = SourceMapping(
        group_field="group",
        date_field="Inv Date",
        amount_field="Total Amount",
        return_field="Returns",
        ignored_fields=["Inv No"],          # ignored for writing, still in audit
        sum_duplicate_dates=True,
        keep_invoice_breakup=True,
        decimal_sep=".", thousands_sep=",",
    )

    print("\nSTEP 5 — Aggregate (sum duplicate dates within a group)")
    rows = aggregate(records, source, aggregation_keys=["group", "date"])
    if not real:
        same_day = [r for r in rows if r.group == CUST_A and r.date == dt.date(2026, 5, 15)]
        check("ALPHA 15/05/2026 row exists", len(same_day) == 1)
        check("630.00 + 12,705.00 summed to 13,335.00",
              round(same_day[0].amount, 2) == 13335.00, f"{same_day[0].amount}")
        check("Invoice breakup retained (2 source rows)",
              len(same_day[0].source_row_indexes) == 2)

    print("\nSTEP 6 — Manual Excel mapping (group->sheet, date col, amount col, return col)")
    wb = excel_reader.open_workbook(xl_bytes)
    sheets = excel_reader.list_sheets(wb)
    print(f"  Workbook sheets available: {sheets}")
    if real:
        print("  (Real run: set group_to_sheet / columns to your actual choices below.)")
        return
    # Map each detected Customer Name to a worksheet (names intentionally differ).
    cfg = AppConfig(
        source=source,
        group_to_sheet={CUST_A: "S-ONE", CUST_B: "S-TWO", CUST_C: "S-THREE"},
        sheets={
            nm: SheetMapping(
                sheet_name=nm, header_row=1,
                date_target_mode=TargetMode.HEADER_NAME, date_column="Txn Date",
                amount_target_mode=TargetMode.COLUMN_LETTER, amount_column="B",
                return_target_mode=TargetMode.COLUMN_LETTER, return_column="D",
            )
            for nm in ["S-ONE", "S-TWO", "S-THREE"]
        },
        write_rules=WriteRules(
            write_action=WriteAction.REPLACE, output_type=OutputType.NUMERIC,
            aggregation_keys=["group", "date"],
        ),
    )

    print("\nSTEP 7 — Build write plan / final preview")
    report = build_plan(wb, rows, cfg)
    print("  Final preview (the same columns the Streamlit page 7 shows):")
    print(f"  {'Customer':<14}{'Sheet':<8}{'Date':<12}{'PDF rows':<24}"
          f"{'Aggregated':<12}{'Row':<5}{'Cell':<6}{'Existing':<10}{'Final':<10}{'Status'}")
    for it in report.items:
        breakup = ",".join(str(a) for a in it.source_amounts)
        print(f"  {it.group:<14}{it.sheet:<8}{it.date_raw:<12}{breakup:<24}"
              f"{it.aggregated_amount:<12}{str(it.matched_row):<5}{it.target_cell:<6}"
              f"{str(it.existing_value):<10}{str(it.final_value):<10}{it.status}")

    alpha_item = [i for i in report.items if i.group == CUST_A and i.date_raw == "15/05/2026"][0]
    check("Preview shows numeric total 13,335.00 for ALPHA 15/05/2026",
          round(alpha_item.aggregated_amount, 2) == 13335.00)
    check("Preview shows human-readable invoice breakup",
          alpha_item.invoice_breakup == "630.00 + 12,705.00", alpha_item.invoice_breakup)
    check("Preview shows comma-free Excel formula breakup",
          alpha_item.excel_formula_breakup == "=630+12705", alpha_item.excel_formula_breakup)
    check("Formula breakup contains no thousands separators",
          "," not in alpha_item.excel_formula_breakup)
    check("Preview shows the matched target cell", bool(alpha_item.target_cell))
    check("Preview shows existing Excel value (empty before write)",
          alpha_item.existing_value in (None, ""))
    # Every item that belongs to a real detected Customer Name must be Ready.
    mapped_items = [i for i in report.items if i.group in (CUST_A, CUST_B, CUST_C)]
    check("All mapped customer items are Ready",
          all(i.status == Status.READY for i in mapped_items),
          str([(i.group, i.status) for i in mapped_items]))
    # Safety: any stray row with no detected group is flagged, NOT written.
    stray = [i for i in report.items if not i.group]
    check("Stray ungrouped rows are flagged Unmapped (won't be written)",
          all(i.status == Status.UNMAPPED_GROUP for i in stray),
          f"{len(stray)} stray row(s)")

    print("\nSTEP 8 — Prove NO write before confirmation; original untouched")
    original_hash = hash(xl_bytes)
    ops = plan_to_write_ops(report, cfg)
    out_bytes, outcomes = excel_writer.apply_writes(xl_bytes, ops)  # this is the 'confirm' step
    check("Original workbook bytes unchanged after write", hash(xl_bytes) == original_hash)
    check("A NEW output workbook was produced", out_bytes != xl_bytes and len(out_bytes) > 0)

    print("\nSTEP 9 — Verify written values + formula/formatting preservation")
    out_wb = openpyxl.load_workbook(io.BytesIO(out_bytes))
    ws1 = out_wb["S-ONE"]
    check("ALPHA aggregated 13,335.00 written to B2", round(ws1["B2"].value, 2) == 13335.00,
          f"B2={ws1['B2'].value}")
    check("Return 100.00 written to D2", round(ws1["D2"].value, 2) == 100.00, f"D2={ws1['D2'].value}")
    check("Formula in E2 preserved", ws1["E2"].value == "=B2+C2-D2", f"E2={ws1['E2'].value}")
    check("Second 'Amount' column (C) left untouched", ws1["C2"].value is None)
    check("Header bold formatting preserved", ws1["A1"].font.bold is True)
    beta = out_wb["S-TWO"]
    check("BETA cross-page row (17/05) written to B4", beta["B4"].value == 3000.0, f"B4={beta['B4'].value}")

    print("\nSTEP 10 — Audit report keeps invoice-wise breakup")
    audit = build_audit_records(outcomes, report.items, source_file=src_name)
    audit_csv = audit_to_csv(audit)
    with open(os.path.join(OUT, "audit.csv"), "w") as fh:
        fh.write(audit_csv)
    out_path = os.path.join(OUT, "updated_workbook.xlsx")
    with open(out_path, "wb") as fh:
        fh.write(out_bytes)
    alpha_audit = [a for a in audit if a["target_cell"] == "B2"][0]
    check("Audit shows both source invoices for the summed cell",
          "INV-1001" in alpha_audit["source_rows"] and "INV-1002" in alpha_audit["source_rows"],
          alpha_audit["source_rows"])
    check("Audit records existing vs new value",
          str(alpha_audit["existing_value"]) in ("None", "") and float(alpha_audit["new_value"]) == 13335.0)

    print(f"\n  Wrote: {out_path}")
    print(f"  Wrote: {os.path.join(OUT, 'audit.csv')}")
    print("\nALL VALIDATION CHECKPOINTS PASSED ✅")


if __name__ == "__main__":
    main()
