# Report → Excel workbook updater

Updates an **existing** Excel workbook from a grouped PDF (or scanned/image)
report — for example a *Customer wise sales details* report whose invoices are
summed per customer per day and written into each customer's worksheet.

The mapping (which customer goes to which worksheet, which columns to use) is
set up **once**, saved as a local profile, and recognised automatically from
the report's structure every time after that.

## Monthly use

```bash
pip install -r requirements.txt
streamlit run app.py
```

1. **Upload** the report (PDF or image) and the workbook to update.
2. **Review.** The app shows *“Previous mapping recognised and loaded”*, a
   check that the extracted invoices match every total printed in the report,
   the customer → worksheet summary, and a preview of every cell it will write
   (customer, date, invoice count and breakup, total, worksheet, cell, existing
   value, new value, status). Only problems are highlighted.
3. **Update.** Tick the confirmation and click **Update uploaded workbook**,
   then download `Updated_<your file name>.xlsx`. An audit (CSV/JSON) and the
   untouched original are offered too.

The button stays disabled when nothing is ready, and the app says *“Nothing is
ready to write.”* It never reports success for zero changes.

## First use (or when something changes)

Setting up takes two steps, both remembered:

1. **Customer → Worksheet.** Choose the worksheet for each customer (or ignore
   a customer). Names are matched later ignoring case, punctuation and spacing,
   so `NORTH MARKET` also finds `NORTHMARKET` — but never by guessing between
   two candidates. Mappings saved for other report types are pre-filled.
2. **Worksheet pattern.** The app takes the pattern from the mapped worksheet
   whose headers most other mapped worksheets share (never an unrelated sheet
   such as a consolidated summary), detects the DATE column and — when a header
   matches the report's Returns field — the returns column. You choose once
   which column the report's main amount goes to. The dropdowns always show the
   headers of the worksheet currently selected.

The pattern is stored by **header name**, so it applies automatically to every
mapped worksheet that has those headers, even if a column or the header row
has moved. Only worksheets laid out differently appear under **Needs
attention**, each with its own small editor; nothing else has to be configured
sheet by sheet.

On later runs the Review step only shows:

```
Customer → Worksheet     (table: worksheet, dates found, target, status)
Worksheet pattern        DATE → A / DATE
                         Total Amount → B / CHALLAN
                         Returns → C / RETURN
                         Applied automatically to 12 of 13 mapped worksheets
```

Everything technical (report fields, all customers, the pattern, per-sheet
exceptions, write settings) is under **Advanced mapping**.

## What it guarantees

* **Only genuine invoice rows are used.** Columns are found from the position
  of the report's header (ruled header boxes or word gaps), and every row must
  have a real date and a numeric amount. Titles, From/To lines, repeated
  headers, printed totals and page footers are excluded, and an invoice number
  can never be taken as a date. Sections that continue onto the next page stay
  with their customer.
* **Reconciliation before writing.** The per-customer and overall totals
  printed in the report must equal the sum of the extracted invoices, or the
  update is blocked with the discrepancy shown.
* **Fields by meaning, not position.** A saved profile stores field names
  (e.g. *Inv Date*, *Total Amount*). On load, the same labelled column must
  exist and contain the right kind of data; if not, only that field is flagged
  for re-mapping. Another column is never substituted.
* **Safe customer lookup.** Names match ignoring case, punctuation and spacing
  (`NORTH MARKET` = `NORTHMARKET`) only when unambiguous; no fuzzy guessing.
* **Blank is not zero.** A blank Returns value in the report never writes 0
  into a RETURN cell (returns may come from a separate report). Zero-filling is
  an explicit per-profile option.
* **The same workbook, changed only where approved.** The uploaded file is
  copied part-for-part and only the target cells' values are changed. Fonts,
  alignment, borders, fills, number formats, comments, widths, heights, views,
  print settings, merged cells and formulas are untouched. Excel recalculates
  dependent formulas on open. After writing, the app proves every other part is
  unchanged before offering the download. Formula cells are never overwritten.

## Profiles

Profiles are stored locally, **outside the repository**, in
`~/.pdf-image-excel/profiles` (Windows: `%APPDATA%\pdf-image-excel\profiles`;
override with `PDFX_PROFILE_DIR`). They contain your business mapping and must
not be committed. The sidebar offers **Use last successful mapping, Import /
Export profile JSON, Rename, Reset mapping and Delete**. A mapping JSON from the
earlier version of this app can be imported; its customer→worksheet and column
settings are kept, and the report fields are re-confirmed by name.

`config/example_profile.json` shows the format with placeholder names.

*Advanced settings* (collapsed) hold the full mapping editor, numeric vs
formula-breakup output (`=12500+7250`), overwrite vs keep existing values,
returns options, date order (day-first by default), reconciliation details,
rejected lines, and CSV exports.

## Tests

```bash
pytest
```

The suite builds realistic PDFs and workbooks with neutral names and covers
parsing, reconciliation, profiles, the in-place workbook update and the app's
three-stage flow.

### Real-file check (local only)

Keep real files, profiles and expected figures in `local/` (gitignored), then:

```bash
python scripts/verify_real_files.py --pdf "local/report.pdf" --xlsx "local/book.xlsx" \
    --expect local/expect.json [--exclude-blocked]
# or
PDFX_REAL_PDF=local/report.pdf PDFX_REAL_XLSX=local/book.xlsx pytest tests/test_real_files.py -s
```

It runs the same pipeline as the app, updates the workbook, and re-checks every
cell of every sheet (values, formulas, formatting), that each customer wrote
only to its own worksheet on the matching date row, and that nothing else
changed.

## Layout

```
app.py                  three-stage Streamlit app
src/values.py           strict amount/date parsing, matching keys
src/extractor.py        PDF/image -> positioned words (pdfplumber, PyMuPDF, OCR)
src/parser.py           groups, columns, invoices, reconciliation
src/profiles.py         profile schema, local store, safe customer lookup
src/workbook.py         read workbook; update target cells in place; verify
src/plan.py             daily totals and the cell-by-cell write plan
src/pipeline.py         prepare / update, shared by app and scripts
src/audit.py            audit and export tables
scripts/verify_real_files.py
```

Scanned PDFs and images need the [Tesseract](https://github.com/tesseract-ocr/tesseract)
program installed; OCR confidence is reported and values should be checked in
the preview.
