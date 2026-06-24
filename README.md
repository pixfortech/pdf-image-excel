# PDF / Image → Excel Smart Mapper

A generic, reusable Streamlit app that extracts tabular/grouped data from a
**PDF or image**, lets you **map** the extracted fields to an **existing Excel
workbook**, previews exactly what will be written, and — only after you confirm
— writes the approved values into a **new copy** of the workbook **without
damaging formatting, formulas, styles, merged cells, or existing structure**.

> **Nothing is hardcoded.** No customer name, branch, worksheet name, column
> name/letter, header text, date format, invoice layout, or business rule is
> baked into the code. Everything is driven by your UI choices and saved JSON
> mappings, so the same app works for completely different PDFs and Excel files.

---

## 1. What the app does

1. Reads a PDF (digital or scanned) or an image and extracts text, tables,
   groups, dates, amounts and row-level records.
2. Detects **group sections** (e.g. a line like `Customer Name: <something>`)
   and assigns the rows beneath each group to that group — continuing the same
   group across page breaks until the next group appears.
3. Shows the raw extracted data in an **editable preview** so you can correct
   values or ignore rows.
4. Lets you **map** which detected field is the group, the date, the amount and
   the optional return/deduction.
5. Lets you **map each detected group to any worksheet** in your Excel file
   (group name and sheet name need not match).
6. Lets you choose the Excel **date column** and **target columns/cells**, by
   header name, column letter, or exact cell reference.
7. **Groups and sums** duplicate dates (optional, configurable aggregation key).
8. Shows a **final write plan** — source group, sheet, date, matched row, target
   cell, existing value, final value, action and status.
9. Writes only the approved cells into a **new workbook**, after you confirm,
   and produces **audit / export files**.

## 2. Install

```bash
pip install -r requirements.txt
```

For scanned PDFs and images you also need the Tesseract binary on your system
for `pytesseract`:

* Ubuntu/Debian: `sudo apt-get install tesseract-ocr`
* macOS: `brew install tesseract`
* Windows: install from the Tesseract project and add it to your PATH.

(Or install `easyocr` instead — see the commented line in `requirements.txt`.)
The app degrades gracefully and tells you if no OCR engine is available.

## 3. Run

```bash
streamlit run app.py
```

Then open the URL Streamlit prints (usually http://localhost:8501).

## 4. Upload PDF/image and Excel

On **page 1 (Upload)**:

* Upload a PDF or image (`PDF, JPG, JPEG, PNG, WEBP, BMP, TIFF`).
* Upload your existing `.xlsx` workbook **or** paste a direct downloadable link
  to it.
* Choose an extraction mode: **Auto** (text first, OCR fallback), **PDF text**,
  **OCR**, or **Table detection**.
* Click **Start extraction**.

## 5. Extract and preview data

On **page 2 (Extraction Preview)**:

* Pick (or confirm the auto-detected) **group label** — the text that
  introduces a group section in your document.
* Choose the parsing approach: detected tables, text lines, or a custom regex.
* Click **Parse rows**.
* Review the raw data in the editable grid. Correct any group, date, or amount;
  tick **ignored** to drop a row. OCR confidence and warnings are shown.

## 6. Map source fields

On **page 3 (Source Field Mapping)** choose which detected field is the:

* Group / customer / branch / account
* Date
* Main amount
* Return / deduction (optional)
* Ignored fields

…and set parsing rules (thousands/decimal separators, currency symbols, date
formats) and whether to **sum duplicate dates**.

### Date handling (format-independent matching)

The app never depends on one fixed date pattern. PDF dates and Excel dates are
each parsed from many representations and **normalised to a plain date**
(`YYYY-MM-DD`) before matching, so different formats still match:

* PDF text: `15/05/2026`, `15-05-2026`, `15.05.2026`, `15 May 2026`,
  `15-May-2026`, `2026-05-15`, `15/05/26`.
* Excel cells: real date objects, datetimes, text dates, `DD/MM/YYYY`,
  `DD-MM-YYYY`, `YYYY-MM-DD`, `DD MMM YYYY`, **Excel serial numbers**, and
  values like `2026-05-15 00:00:00`.

A **Date interpretation** setting disambiguates numeric dates like `05/06/2026`:

* **Indian / British — DD/MM/YYYY (default)** → 5 June 2026
* **US — MM/DD/YYYY** → 6 May 2026… i.e. May 6
* **Auto-detect**

**Excel dates are read by their real stored value, never by display
formatting.** A cell formatted `d-mmm` shows `04-Jan` but stores a real
`datetime(2026, 1, 4)`; openpyxl returns the true datetime, so the year is taken
from the value, not the display. The full date column is scanned (not just a
preview window).

Page 7 shows a date-matching analysis per mapped sheet: selected date column,
the **first 10 and last 10** parsed Excel dates (raw → normalised), the
**minimum and maximum** Excel date (full-column scan), the PDF date range,
matched / missing counts, and **suggested alternative date columns**. A
**`excel_date_debug.csv`** export lists every date cell's `raw_cell_value`,
`cell_data_type`, `number_format` and `parsed_date` so a formatting issue is
distinguishable from a wrong-year/missing-row issue.

If more than 25% of PDF dates are missing, a **strong warning** explains the
likely cause and shows both date ranges. If a PDF date falls **within** the
Excel column's date range but no exact row exists (or a date present in the
column fails to match — a genuine bug), the row is flagged as an **error** so
nothing is written until it is resolved.

Missing dates are **never inserted silently** — the default action is **skip**;
rows are inserted/copied only if you explicitly choose that behaviour on page 6.

## 7. Map Excel sheets and columns

First, on **page 5 (Group → Sheet Mapping)** assign each detected Customer Name
group to a worksheet (or skip it). Group name and worksheet name are
independent — e.g. `BARANAGAR → BN`, `BEADON STREET → BD`, `KESTOPUR → KP`.

Then, on **page 4 (Excel Mapping)** you do **not** have to configure every sheet
by hand. Most branch sheets share the same layout, so the page works as a
**mapping pattern (template)**:

* **A. Create mapping pattern from this worksheet** — pick one sample worksheet
  and define its **header row**, **date column**, **amount target column**,
  **return target column**, plus the **write mode**, **existing-value
  behaviour** and **missing-date behaviour**.
* **B. Apply this column/date pattern to all mapped worksheets** — one click
  copies that column pattern to every worksheet assigned to a group. Each copy
  keeps its own sheet name; only the column layout is shared. **Column letter is
  preferred** for copied patterns because it is reliable when headers repeat
  (exact cell references are sheet-specific and are dropped on copies).
* **C. Per-sheet mapping status** — a table showing, per Customer Name: assigned
  sheet, header row, date column, amount/return columns and a **status**
  (`Ready`, `Sheet missing`, `Date column missing`, `Amount column missing`,
  `Return column missing`, `Needs review`).
* **D. Override an individual worksheet** — if one sheet has a different
  structure, override just that sheet; the others keep the pattern.

Targets can always be selected **by header name, column letter, or exact cell
reference**, which matters when a sheet has **repeated header names**.

If several groups are mapped but only one worksheet has a column mapping, the
app warns: *"Only one worksheet has column mapping. Apply this pattern to all
mapped worksheets or configure each worksheet before writing."*

The final write preview uses the worksheet **assigned to each Customer Name**
with that sheet's own column mapping — it never writes every group into the one
template worksheet.

On **page 6 (Aggregation & Write Rules)** choose aggregation keys, sum behaviour,
what to do with existing cell values (**replace / add / skip / ask**), the
**output mode** (see below), and what to do when a date row is missing.

### Output mode: numeric total vs Excel formula breakup

For every aggregated group/date the app keeps **three** value forms, and the
numeric total is **always** preserved (never replaced) for audit and accounting:

| Field | Example | Purpose |
|-------|---------|---------|
| `aggregated_amount` | `13335.00` | clean numeric total — safe for accounting, CSV, audit |
| `invoice_breakup` | `630.00 + 12,705.00` | human-readable preview (plain text, **not** a formula) |
| `excel_formula_breakup` | `=630+12705` | a real Excel formula — starts with `=`, **no** thousands separators |

The **Output mode** radio on page 6 decides what is written into the target cell:

* **Numeric total** — writes the numeric `aggregated_amount` (e.g. `13335.0`).
  Recommended for the actual accounting sheet: cleaner and safer.
* **Excel formula breakup** — writes the `excel_formula_breakup` (e.g.
  `=630+12705`) so Excel shows and recalculates the breakup.

All three values appear in the final write-plan preview, in `write_plan.csv`,
in `grouped_totals.csv`, and in the audit report — regardless of which output
mode you pick.

## 8. Confirm before writing

On **page 7 (Final Preview & Validation)** the app shows the full write plan and
a validation summary (unmapped groups, dates not found, missing columns,
conflicts, duplicate headers). The write button is **disabled** until blocking
issues are resolved **and** you tick the confirmation box. **Nothing is written
to Excel before you confirm.**

## 9. Download the updated workbook

On **page 8 (Export)** download:

* The **updated Excel workbook** (a new file; your original is never modified)
* A **backup** of the original
* **Audit report** as CSV and JSON (includes numeric total, invoice breakup and Excel formula breakup)
* **Extracted raw data** CSV
* **Grouped / aggregated data** CSV (`grouped_totals.csv` — numeric total + breakup + formula)
* **Write plan** CSV (`write_plan.csv` — per target cell, all three value forms + write mode)
* The **saved mapping** JSON
* An **error / warning report** CSV

## 10. Reuse saved mapping JSON

Use the sidebar **Download current mapping JSON** to save your configuration.
Next time, use **Load Previous Mapping** in the sidebar to restore the same
source mapping, group→sheet assignments, Excel column mapping and write rules —
so you don't have to remap a recurring PDF/Excel format. Sample (placeholder)
mapping files live in `config/`.

## 11. Run tests

```bash
pip install -r requirements.txt
pytest
```

The tests cover grouped-row extraction, group continuation across page breaks,
ignoring repeated headers/totals, comma amount parsing, date parsing, summing
duplicate group/date rows, non-hardcoded group→sheet mapping, column selection
by letter, repeated Excel headers, formula preservation, no-overwrite-without-
confirmation, audit generation, the **numeric-total vs Excel-formula-breakup**
behaviour (`630.00 + 12,705.00` → numeric `13335.00` and formula `=630+12705`,
with commas stripped from formulas), appending multiple missing dates to
distinct rows, and the **Excel mapping-pattern** workflow (copying one
worksheet's column pattern to many mapped sheets, per-sheet overrides, and
verifying each Customer Name writes to its own assigned sheet rather than the
template sheet), and **format-independent date handling** (DD/MM/YYYY,
DD-MM-YYYY, YYYY-MM-DD, datetimes, Excel serial dates, text and two-digit-year
dates all normalised; Indian/British default interpretation; rejecting non-date
text like `wise`/`Date:`; the strong warning when most dates are missing).

---

## Project structure

```
pdf-image-excel/
  app.py                      Streamlit UI (8 workflow pages)
  requirements.txt
  README.md
  config/
    sample_pdf_mapping.json   placeholder examples (no real business values)
    sample_excel_mapping.json
    sample_write_rules.json
  src/
    extractor.py     PDF text/table extraction (pdfplumber → PyMuPDF → OCR)
    ocr.py           pytesseract / EasyOCR with confidence reporting
    parser.py        group detection, page continuation, header/total filtering
    table_detector.py table normalisation helpers
    mapping.py       dataclass config + JSON save/load (no hardcoded rules)
    aggregator.py    configurable grouping & summing
    excel_reader.py  sheet/column/cell resolution, date-row matching
    excel_writer.py  safe writes preserving formatting & formulas
    validator.py     write-plan build + validation statuses
    audit.py         CSV/JSON audit and export generation
    utils.py         amount & date parsing primitives
  tests/
    test_extractor.py  test_parser.py  test_mapping.py
    test_aggregator.py test_excel_writer.py test_excel_mapping.py
    test_formula.py    test_dates.py     test_excel_dates.py
    test_end_to_end.py
  scripts/
    real_file_validation.py   end-to-end harness (generates real PDF + .xlsx,
                              or accepts your real file paths as arguments)
```

## Design notes on safety & "no hardcoding"

* The uploaded original file is **never** modified in place. Writing loads a
  fresh workbook from the original bytes, edits only approved target cells, and
  saves a **new** file. A separate backup copy is also offered.
* openpyxl preserves untouched cells, formulas, styles, merged cells, column
  widths, row heights, number formats, borders and fills; the app only assigns
  `.value` on explicitly approved targets.
* All field names, labels, customer/group names, worksheet names, columns, date
  formats and write rules come from the user (UI or loaded JSON). The example
  names mentioned in the task brief (branches, headers, etc.) appear **only** in
  documentation/tests as illustrations — never in application logic.
