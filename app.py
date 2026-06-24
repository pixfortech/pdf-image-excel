"""Smart PDF / Image to Excel data extraction and mapping app (Streamlit UI).

This is a generic, reusable tool: it makes NO assumptions about customer names,
worksheet names, columns, headers, dates, amounts or document layouts.  Every
business rule is supplied by the user through the UI and/or a saved JSON
configuration.

Run with:
    streamlit run app.py
"""
from __future__ import annotations

import datetime as _dt
import io
import json

import pandas as pd
import streamlit as st

from src import audit as audit_mod
from src import excel_reader, excel_writer, extractor, mapping as mapping_mod, parser
from src.aggregator import aggregate, grouped_export_rows
from src.mapping import (
    AppConfig,
    DateNotFoundAction,
    OutputType,
    SheetMapping,
    SourceMapping,
    TargetMode,
    WriteAction,
    WriteRules,
    config_from_json,
    config_to_json,
)
from src.validator import (
    Status, build_plan, configured_mapped_sheets, plan_to_export_rows,
    plan_to_write_ops, sheet_mapping_status,
)

st.set_page_config(page_title="PDF/Image → Excel Mapper", layout="wide")


# ---------------------------------------------------------------------------
# Session state helpers
# ---------------------------------------------------------------------------

def init_state():
    ss = st.session_state
    ss.setdefault("config", AppConfig())
    ss.setdefault("extraction", None)       # ExtractionResult
    ss.setdefault("records", None)          # list[Record]
    ss.setdefault("records_df", None)       # editable DataFrame
    ss.setdefault("excel_bytes", None)      # original workbook bytes
    ss.setdefault("excel_name", "")
    ss.setdefault("source_name", "")
    ss.setdefault("plan", None)
    ss.setdefault("output_bytes", None)
    ss.setdefault("outcomes", None)
    ss.setdefault("audit_records", None)


init_state()
CFG: AppConfig = st.session_state["config"]


def workbook():
    if st.session_state["excel_bytes"] is None:
        return None
    return excel_reader.open_workbook(st.session_state["excel_bytes"])


# ---------------------------------------------------------------------------
# Record <-> DataFrame conversion and small UI helpers
# ---------------------------------------------------------------------------

_META_COLS = ["page", "group", "source_text", "ocr_confidence", "status", "warnings", "ignored", "row_index"]


def records_to_df(records):
    """Flatten parser Records into an editable DataFrame (one row per record)."""
    rows = []
    for r in records:
        row = {
            "row_index": r.row_index,
            "page": r.page,
            "group": r.group,
            "ignored": r.ignored,
            "status": r.status,
            "ocr_confidence": r.ocr_confidence,
            "warnings": "; ".join(r.warnings),
            "source_text": r.source_text,
        }
        row.update(r.fields)
        rows.append(row)
    df = pd.DataFrame(rows)
    # Order: row_index, page, group, ignored, then detected fields, then meta.
    field_cols = [c for c in df.columns if c not in _META_COLS]
    ordered = ["row_index", "page", "group", "ignored"] + field_cols + \
              ["status", "ocr_confidence", "warnings", "source_text"]
    return df[[c for c in ordered if c in df.columns]]


def df_to_records(df):
    """Rebuild parser Records from the (possibly edited) DataFrame."""
    from src.parser import Record
    records = []
    field_cols = [c for c in df.columns if c not in _META_COLS]
    for _, row in df.iterrows():
        fields = {c: ("" if pd.isna(row[c]) else str(row[c])) for c in field_cols}
        records.append(Record(
            page=int(row.get("page", 0) or 0),
            group="" if pd.isna(row.get("group")) else str(row.get("group")),
            fields=fields,
            source_text=str(row.get("source_text", "") or ""),
            ocr_confidence=None if pd.isna(row.get("ocr_confidence")) else float(row.get("ocr_confidence")),
            status=str(row.get("status", "ok") or "ok"),
            ignored=bool(row.get("ignored", False)),
            row_index=int(row.get("row_index", -1) or -1),
        ))
    return records


def _select(label, options, current, key=None):
    options = list(options)
    idx = options.index(current) if current in options else 0
    return st.selectbox(label, options, index=idx, key=key)


def _column_selector(label, mode, header_names, letters, current, key=None):
    """Render the right input depending on the selection mode."""
    if mode == TargetMode.HEADER_NAME:
        opts = [""] + header_names
        idx = opts.index(current) if current in opts else 0
        return st.selectbox(f"{label} (by header name)", opts, index=idx, key=key)
    if mode == TargetMode.COLUMN_LETTER:
        opts = [""] + letters
        idx = opts.index(current) if current in opts else 0
        return st.selectbox(f"{label} (by column letter)", opts, index=idx, key=key)
    # cell reference
    return st.text_input(f"{label} (exact cell reference, e.g. C5)", value=current, key=key)


# ---------------------------------------------------------------------------
# Sidebar navigation + config load/save
# ---------------------------------------------------------------------------

PAGES = [
    "1. Upload",
    "2. Extraction Preview",
    "3. Source Field Mapping",
    "4. Excel Mapping",
    "5. Group → Sheet Mapping",
    "6. Aggregation & Write Rules",
    "7. Final Preview & Validation",
    "8. Export",
]

st.sidebar.title("📄 → 📊  Mapper")
page = st.sidebar.radio("Workflow", PAGES)

st.sidebar.markdown("---")
st.sidebar.subheader("Reuse a saved mapping")
loaded = st.sidebar.file_uploader("Load Previous Mapping (JSON)", type=["json"], key="cfg_upload")
if loaded is not None and st.sidebar.button("Apply loaded mapping"):
    try:
        st.session_state["config"] = config_from_json(loaded.read().decode("utf-8"))
        st.sidebar.success("Mapping loaded. Navigate the pages to review it.")
        st.rerun()
    except Exception as exc:
        st.sidebar.error(f"Could not load mapping: {exc}")

st.sidebar.download_button(
    "💾 Download current mapping JSON",
    data=config_to_json(CFG),
    file_name="mapping.json",
    mime="application/json",
)


# ===========================================================================
# PAGE 1 — UPLOAD
# ===========================================================================
if page == PAGES[0]:
    st.header("1. Upload your files")
    st.caption(
        "Upload a PDF or image (the source of the data) and the existing Excel "
        "workbook to write into. Nothing is written until you confirm on page 7."
    )

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Source document")
        src_file = st.file_uploader(
            "PDF or image (PDF, JPG, JPEG, PNG, WEBP, BMP, TIFF)",
            type=["pdf", "jpg", "jpeg", "png", "webp", "bmp", "tiff", "tif"],
            key="src_file",
        )
        mode = st.selectbox(
            "Extraction mode",
            ["auto", "pdf_text", "ocr", "table"],
            help="Auto tries digital text first and falls back to OCR for scanned files.",
        )

    with col2:
        st.subheader("Excel workbook")
        xl_file = st.file_uploader("Existing Excel workbook (.xlsx)", type=["xlsx", "xlsm"], key="xl_file")
        xl_link = st.text_input("…or a direct downloadable Excel link")
        if st.button("Fetch Excel from link") and xl_link:
            try:
                import requests
                resp = requests.get(xl_link, timeout=30)
                resp.raise_for_status()
                st.session_state["excel_bytes"] = resp.content
                st.session_state["excel_name"] = xl_link.split("/")[-1] or "linked.xlsx"
                st.success(f"Downloaded {st.session_state['excel_name']}")
            except Exception as exc:
                st.error(f"Could not download Excel: {exc}")

    if xl_file is not None:
        st.session_state["excel_bytes"] = xl_file.getvalue()
        st.session_state["excel_name"] = xl_file.name

    if st.session_state["excel_bytes"] is not None:
        try:
            sheets = excel_reader.list_sheets(workbook())
            st.info(f"Excel loaded: **{st.session_state['excel_name']}** — sheets: {', '.join(sheets)}")
        except Exception as exc:
            st.error(f"Could not read Excel: {exc}")

    st.markdown("---")
    if st.button("▶️ Start extraction", type="primary"):
        if src_file is None:
            st.error("Please upload a PDF or image first.")
        else:
            with st.spinner("Extracting…"):
                data = src_file.getvalue()
                st.session_state["source_name"] = src_file.name
                result = extractor.extract(data, filename=src_file.name, mode=mode)
                st.session_state["extraction"] = result
                # Auto-suggest a group label from the text.
                cands = parser.detect_group_label_candidates(result.pages_text)
                if cands and not CFG.source.group_label:
                    CFG.source.group_label = cands[0]
            st.success(
                f"Extraction done with engine '{result.engine}'. "
                f"Pages: {result.page_count}. Detected tables: {len(result.tables)}."
            )
            if result.warnings:
                for w in result.warnings:
                    st.warning(w)
            st.info("Go to page **2. Extraction Preview** to review and correct the data.")


# ===========================================================================
# PAGE 2 — EXTRACTION PREVIEW
# ===========================================================================
elif page == PAGES[1]:
    st.header("2. Extraction preview & correction")
    result = st.session_state["extraction"]
    if result is None:
        st.warning("No extraction yet. Go to page 1 and run extraction.")
    else:
        if result.ocr_confidence is not None:
            lvl = "warning" if result.ocr_confidence < 70 else "info"
            getattr(st, lvl)(f"OCR average confidence: {result.ocr_confidence:.1f}%")

        st.subheader("Detected group label")
        cands = parser.detect_group_label_candidates(result.pages_text)
        options = sorted(set([CFG.source.group_label] + cands + [""]))
        CFG.source.group_label = st.selectbox(
            "Which line label introduces a new group? (auto-detected candidates listed)",
            options,
            index=options.index(CFG.source.group_label) if CFG.source.group_label in options else 0,
            help="Examples of such labels could be a 'name' field in your report. Leave blank for auto.",
        )

        st.subheader("Parsing approach")
        approach = st.radio(
            "How should rows be parsed?",
            ["Detected tables (recommended if tables were found)", "Text lines", "Custom regex"],
            horizontal=False,
        )
        regex = ""
        if approach == "Custom regex":
            regex = st.text_input(
                "Row regex (use named groups like (?P<date>...) or positional groups)",
                value="",
            )

        if st.button("🔍 Parse rows", type="primary"):
            if approach.startswith("Detected tables"):
                preferred = "tables"
            elif approach == "Custom regex" and regex:
                preferred = "regex"
            else:
                preferred = "text"
            # Auto-switches tables->text when table mode parses rows but finds 0
            # groups (group labels sitting outside the tables).
            records, mode_used, switched = parser.smart_parse(
                result.tables, result.pages_text,
                group_label=CFG.source.group_label,
                preferred=preferred, row_regex=regex,
            )
            st.session_state["records"] = records
            st.session_state["records_df"] = records_to_df(records)
            n_groups = len(parser.list_detected_groups(records))

            if switched:
                st.warning(
                    "⚠️ Detected-tables mode parsed rows but found **0 groups** — "
                    "the group labels are outside the tables in this PDF. "
                    f"Automatically switched to **Text lines** mode, which found "
                    f"**{n_groups} group(s)**. (You can re-parse with another mode above.)"
                )
            elif mode_used == "tables" and n_groups == 0:
                st.error(
                    "⚠️ Detected-tables mode parsed rows but found **0 groups**, and no "
                    "group markers were found in the page text either. Check the "
                    "**group label** above (page 2) or try **Text lines** / **Custom regex** mode."
                )
            st.success(f"Parsed {len(records)} rows. Detected groups: {n_groups}.")

        df = st.session_state["records_df"]
        if df is not None:
            st.subheader("Raw extracted data (editable)")
            st.caption(
                "Edit any cell to correct values. Toggle 'ignored' to drop a row. "
                "Changes here are used for mapping and writing."
            )
            edited = st.data_editor(df, use_container_width=True, num_rows="fixed", key="editor")
            st.session_state["records_df"] = edited

            groups = sorted({str(g) for g in edited["group"].tolist() if str(g).strip()})
            st.markdown(f"**Detected groups ({len(groups)}):** " + ", ".join(groups) if groups else "**No groups detected yet.**")


# ===========================================================================
# PAGE 3 — SOURCE FIELD MAPPING
# ===========================================================================
elif page == PAGES[2]:
    st.header("3. Source field mapping")
    df = st.session_state["records_df"]
    if df is None:
        st.warning("No parsed rows yet. Complete page 2 first.")
    else:
        field_cols = [c for c in df.columns if c not in ("page", "group", "source_text",
                                                          "ocr_confidence", "status", "warnings", "ignored", "row_index")]
        all_fields = ["group"] + field_cols
        st.caption("Tell the app what each detected field means. Field names come from YOUR document.")

        c1, c2 = st.columns(2)
        with c1:
            CFG.source.group_field = _select("Group / customer / branch / account field",
                                              ["group"] + field_cols, CFG.source.group_field or "group")
            CFG.source.date_field = _select("Date field", [""] + field_cols, CFG.source.date_field)
            CFG.source.amount_field = _select("Main amount field", [""] + field_cols, CFG.source.amount_field)
            CFG.source.return_field = _select("Return / deduction field (optional)", [""] + field_cols, CFG.source.return_field)
        with c2:
            CFG.source.ignored_fields = st.multiselect("Fields to ignore", field_cols, default=CFG.source.ignored_fields)
            CFG.source.sum_duplicate_dates = st.checkbox("Sum duplicate dates within a group", value=CFG.source.sum_duplicate_dates)
            CFG.source.keep_invoice_breakup = st.checkbox("Keep source-row breakup in audit", value=CFG.source.keep_invoice_breakup)

        st.subheader("Parsing rules")
        c3, c4, c5 = st.columns(3)
        with c3:
            CFG.source.thousands_sep = st.text_input("Thousands separator", value=CFG.source.thousands_sep)
        with c4:
            CFG.source.decimal_sep = st.text_input("Decimal separator", value=CFG.source.decimal_sep)
        with c5:
            CFG.source.currency_chars = st.text_input("Currency symbols to strip", value=CFG.source.currency_chars)
        fmts = st.text_input(
            "Date formats (comma-separated strptime patterns; blank = try many)",
            value=", ".join(CFG.source.date_formats),
        )
        CFG.source.date_formats = [f.strip() for f in fmts.split(",") if f.strip()]

        if CFG.source.date_field and CFG.source.amount_field:
            st.success("Source mapping looks complete. Continue to Excel mapping.")
        else:
            st.info("Select at least a date field and an amount field to proceed.")


# ===========================================================================
# PAGE 4 — EXCEL MAPPING
# ===========================================================================
elif page == PAGES[3]:
    st.header("4. Excel sheet & column mapping")
    wb = workbook()
    if wb is None:
        st.warning("No Excel workbook loaded. Go to page 1.")
    else:
        sheets = excel_reader.list_sheets(wb)
        mapped_sheets = [s for s in dict.fromkeys(CFG.group_to_sheet.values()) if s]

        st.caption(
            "Configure ONE worksheet as a column/date pattern, then apply it to all "
            "worksheets assigned to your Customer Name groups. Column letter is "
            "preferred for copied patterns because it is reliable when headers repeat."
        )
        if not mapped_sheets:
            st.info("Tip: assign Customer Names to worksheets on page 5 first, then "
                    "apply the pattern to all of them here.")

        # ---- A. Create / edit a mapping pattern from one worksheet ----
        st.subheader("A. Create mapping pattern from this worksheet")
        template_options = sheets
        default_tmpl = st.session_state.get("template_sheet") or (mapped_sheets[0] if mapped_sheets else sheets[0])
        sheet = st.selectbox("Template / sample worksheet", template_options,
                             index=template_options.index(default_tmpl) if default_tmpl in template_options else 0)
        st.session_state["template_sheet"] = sheet
        sm = CFG.sheets.get(sheet, SheetMapping(sheet_name=sheet))
        sm.sheet_name = sheet

        sm.header_row = st.number_input("Header row number", min_value=1, value=int(sm.header_row or 1))

        with st.expander(f"Preview of '{sheet}'", expanded=True):
            preview = excel_reader.preview_sheet(wb, sheet)
            st.dataframe(pd.DataFrame(preview), use_container_width=True)

        dups = excel_reader.find_duplicate_headers(wb, sheet, sm.header_row)
        if dups:
            st.warning(
                f"Repeated header names in this sheet: {', '.join(dups)}. "
                "Prefer **column letter** or **exact cell reference** for these."
            )

        headers = excel_reader.header_values(wb, sheet, sm.header_row)
        header_names = [h for _, h in headers if h]
        letters = [l for l, _ in headers]
        st.caption("Columns: " + ", ".join(f"{l}={h}" for l, h in headers if h)[:400])

        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("**Date column**")
            sm.date_target_mode = TargetMode(_select("Identify date by",
                                                     [m.value for m in TargetMode], sm.date_target_mode.value, key="dm"))
            sm.date_column = _column_selector("Date column", sm.date_target_mode, header_names, letters, sm.date_column, key="dcol")
        with c2:
            st.markdown("**Amount target**")
            sm.amount_target_mode = TargetMode(_select("Identify amount by",
                                                       [m.value for m in TargetMode], sm.amount_target_mode.value, key="am"))
            if sm.amount_target_mode == TargetMode.CELL_REFERENCE:
                sm.amount_cell = st.text_input("Exact amount cell (e.g. F10)", value=sm.amount_cell)
            else:
                sm.amount_column = _column_selector("Amount column", sm.amount_target_mode, header_names, letters, sm.amount_column, key="acol")
        with c3:
            st.markdown("**Return target (optional)**")
            sm.return_target_mode = TargetMode(_select("Identify return by",
                                                       [m.value for m in TargetMode], sm.return_target_mode.value, key="rm"))
            if sm.return_target_mode == TargetMode.CELL_REFERENCE:
                sm.return_cell = st.text_input("Exact return cell", value=sm.return_cell)
            else:
                sm.return_column = _column_selector("Return column", sm.return_target_mode, header_names, letters, sm.return_column, key="rcol")

        dfmts = st.text_input("Worksheet date formats for text dates (comma-separated, blank = many)",
                              value=", ".join(sm.date_formats), key="sheet_dfmt")
        sm.date_formats = [f.strip() for f in dfmts.split(",") if f.strip()]
        CFG.sheets[sheet] = sm

        st.markdown("**Write behaviour for this pattern** (applies to all sheets):")
        wc1, wc2, wc3 = st.columns(3)
        with wc1:
            _OUT = {OutputType.NUMERIC.value: "Numeric total", OutputType.FORMULA.value: "Excel formula breakup"}
            CFG.write_rules.output_type = OutputType(_select("Write mode",
                                                             [o.value for o in OutputType], CFG.write_rules.output_type.value, key="pat_out"))
        with wc2:
            CFG.write_rules.write_action = WriteAction(_select("Existing value behaviour",
                                                               [a.value for a in WriteAction], CFG.write_rules.write_action.value, key="pat_act"))
        with wc3:
            CFG.write_rules.date_not_found_action = DateNotFoundAction(_select("Missing date behaviour",
                                                                               [a.value for a in DateNotFoundAction], CFG.write_rules.date_not_found_action.value, key="pat_dnf"))

        # ---- B. Apply pattern to all mapped sheets ----
        st.subheader("B. Apply pattern to all mapped worksheets")
        st.caption(f"Mapped worksheets (from page 5): {', '.join(mapped_sheets) if mapped_sheets else '(none yet)'}")
        targets = [s for s in mapped_sheets if s != sheet]
        if st.button("📋 Apply this column/date pattern to all mapped worksheets",
                     type="primary", disabled=not targets):
            written = mapping_mod.apply_pattern_to_sheets(CFG, sheet, mapped_sheets)
            st.success(f"Applied the '{sheet}' pattern to {len(written)} worksheet(s): "
                       f"{', '.join(written) if written else '—'}. Override any individually in section D.")

        # ---- C. Status preview ----
        st.subheader("C. Per-sheet mapping status")
        if mapped_sheets:
            status_rows = sheet_mapping_status(wb, CFG)
            st.dataframe(pd.DataFrame([{
                "Customer Name": r["customer_name"], "Assigned Sheet": r["assigned_sheet"],
                "Header Row": r["header_row"], "Date Column": r["date_column"],
                "Amount Target Column": r["amount_column"], "Return Column": r["return_column"],
                "Status": r["status"],
            } for r in status_rows]), use_container_width=True)

            configured = configured_mapped_sheets(CFG)
            if len(mapped_sheets) > 1 and len(configured) <= 1:
                st.warning(
                    "Only one worksheet has column mapping. Apply this pattern to all "
                    "mapped worksheets or configure each worksheet before writing."
                )
        else:
            st.info("No groups are mapped to worksheets yet (do that on page 5).")

        # ---- D. Per-sheet override ----
        st.subheader("D. Override an individual worksheet (optional)")
        if mapped_sheets:
            ov_sheet = st.selectbox("Override mapping for", mapped_sheets, key="override_sheet")
            if ov_sheet and ov_sheet != sheet:
                osm = CFG.sheets.get(ov_sheet, SheetMapping(sheet_name=ov_sheet))
                osm.sheet_name = ov_sheet
                osm.header_row = st.number_input("Header row", min_value=1, value=int(osm.header_row or 1), key="ov_hr")
                oheaders = excel_reader.header_values(wb, ov_sheet, osm.header_row)
                ohn = [h for _, h in oheaders if h]
                oletters = [l for l, _ in oheaders]
                oc1, oc2, oc3 = st.columns(3)
                with oc1:
                    osm.date_target_mode = TargetMode(_select("Date by", [m.value for m in TargetMode], osm.date_target_mode.value, key="ov_dm"))
                    osm.date_column = _column_selector("Date column", osm.date_target_mode, ohn, oletters, osm.date_column, key="ov_dcol")
                with oc2:
                    osm.amount_target_mode = TargetMode(_select("Amount by", [m.value for m in TargetMode], osm.amount_target_mode.value, key="ov_am"))
                    if osm.amount_target_mode == TargetMode.CELL_REFERENCE:
                        osm.amount_cell = st.text_input("Amount cell", value=osm.amount_cell, key="ov_acell")
                    else:
                        osm.amount_column = _column_selector("Amount column", osm.amount_target_mode, ohn, oletters, osm.amount_column, key="ov_acol")
                with oc3:
                    osm.return_target_mode = TargetMode(_select("Return by", [m.value for m in TargetMode], osm.return_target_mode.value, key="ov_rm"))
                    if osm.return_target_mode == TargetMode.CELL_REFERENCE:
                        osm.return_cell = st.text_input("Return cell", value=osm.return_cell, key="ov_rcell")
                    else:
                        osm.return_column = _column_selector("Return column", osm.return_target_mode, ohn, oletters, osm.return_column, key="ov_rcol")
                CFG.sheets[ov_sheet] = osm
                st.caption(f"Override saved for '{ov_sheet}'. This sheet no longer follows the pattern.")
            elif ov_sheet == sheet:
                st.caption("This is the template worksheet — edit it in section A above.")


# ===========================================================================
# PAGE 5 — GROUP → SHEET MAPPING
# ===========================================================================
elif page == PAGES[4]:
    st.header("5. Group → worksheet mapping")
    df = st.session_state["records_df"]
    wb = workbook()
    if df is None or wb is None:
        st.warning("Need parsed rows (page 2) and an Excel workbook (page 1).")
    else:
        groups = sorted({str(g) for g in df["group"].tolist() if str(g).strip()})
        sheets = ["(skip)"] + excel_reader.list_sheets(wb)
        st.caption(
            "Assign each detected group to a worksheet. Group name and worksheet "
            "name need NOT match. Leave as '(skip)' to exclude a group."
        )
        for g in groups:
            current = CFG.group_to_sheet.get(g, "(skip)")
            idx = sheets.index(current) if current in sheets else 0
            choice = st.selectbox(f"Group: {g}", sheets, index=idx, key=f"g2s_{g}")
            if choice == "(skip)":
                CFG.group_to_sheet.pop(g, None)
            else:
                CFG.group_to_sheet[g] = choice
        st.success("Assignments saved. They are stored in the mapping JSON (sidebar).")


# ===========================================================================
# PAGE 6 — AGGREGATION & WRITE RULES
# ===========================================================================
elif page == PAGES[5]:
    st.header("6. Aggregation & write rules")
    wr = CFG.write_rules
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Aggregation")
        key_opts = ["group", "date"]
        if CFG.source.amount_field:
            key_opts += [c for c in (st.session_state["records_df"].columns if st.session_state["records_df"] is not None else [])
                         if c not in ("page", "group", "source_text", "ocr_confidence", "status", "warnings", "ignored", "row_index")]
        wr.aggregation_keys = st.multiselect("Aggregation keys", key_opts, default=wr.aggregation_keys)
        CFG.source.sum_duplicate_dates = st.checkbox("Sum duplicate dates", value=CFG.source.sum_duplicate_dates)
    with c2:
        st.subheader("Write behaviour")
        wr.write_action = WriteAction(_select("On existing cell value",
                                              [a.value for a in WriteAction], wr.write_action.value))
        # Output mode: numeric total vs Excel formula breakup.
        _OUT_LABELS = {OutputType.NUMERIC.value: "Numeric total (recommended for accounting)",
                       OutputType.FORMULA.value: "Excel formula breakup (=630+12705)"}
        out_choice = st.radio(
            "Output mode (what lands in the target cell)",
            [OutputType.NUMERIC.value, OutputType.FORMULA.value],
            index=[OutputType.NUMERIC.value, OutputType.FORMULA.value].index(wr.output_type.value),
            format_func=lambda v: _OUT_LABELS[v],
        )
        wr.output_type = OutputType(out_choice)
        st.caption(
            "The numeric total is always kept for audit/CSV regardless of this choice. "
            "Numeric mode writes e.g. 13335.0; formula mode writes e.g. =630+12705 so "
            "Excel shows the breakup and recalculates."
        )
        wr.add_source_comment = st.checkbox("Add a cell comment with source rows", value=wr.add_source_comment)

    st.subheader("When the date is not found in the worksheet")
    wr.date_not_found_action = DateNotFoundAction(_select(
        "Action", [a.value for a in DateNotFoundAction], wr.date_not_found_action.value))
    wr.insert_missing_date_rows = st.checkbox("Insert a new row when date missing", value=wr.insert_missing_date_rows)
    st.info("Default safety: existing values are never overwritten unless you choose 'replace' or 'add'.")


# ===========================================================================
# PAGE 7 — FINAL PREVIEW & VALIDATION
# ===========================================================================
elif page == PAGES[6]:
    st.header("7. Final preview & validation")
    df = st.session_state["records_df"]
    wb = workbook()
    if df is None or wb is None:
        st.warning("Need parsed rows and an Excel workbook.")
    elif not (CFG.source.date_field and CFG.source.amount_field):
        st.warning("Complete source field mapping (page 3) first.")
    else:
        records = df_to_records(df)
        rows = aggregate(records, CFG.source, aggregation_keys=CFG.write_rules.aggregation_keys)
        report = build_plan(wb, rows, CFG)
        st.session_state["plan"] = report

        # Warn if many groups are mapped but only one sheet has a column mapping.
        mapped_sheets = [s for s in dict.fromkeys(CFG.group_to_sheet.values()) if s]
        configured = configured_mapped_sheets(CFG)
        if len(mapped_sheets) > 1 and len(configured) <= 1:
            st.warning(
                "Only one worksheet has column mapping. Apply this pattern to all "
                "mapped worksheets or configure each worksheet before writing (page 4)."
            )

        st.subheader("Validation summary")
        cols = st.columns(5)
        cols[0].metric("Ready", report.ready_count)
        cols[1].metric("Unmapped groups", len(report.unmapped_groups))
        cols[2].metric("Dates not found", report.dates_not_found)
        cols[3].metric("Missing columns", report.missing_columns)
        cols[4].metric("Conflicts", report.conflicts)

        if report.unmapped_groups:
            st.warning("Unmapped groups: " + ", ".join(report.unmapped_groups))
        if report.duplicate_headers:
            for sh, d in report.duplicate_headers.items():
                st.warning(f"Sheet '{sh}' has duplicate headers: {', '.join(d)}")

        write_mode = CFG.write_rules.output_type.value
        st.subheader(f"Write plan (exactly what will happen) — write mode: {write_mode}")
        st.caption("All three value forms are shown. Only 'Value to write' lands in the cell.")
        plan_rows = []
        for it in report.items:
            plan_rows.append({
                "Customer Name": it.group,
                "Worksheet": it.sheet,
                "Date": it.date_raw,
                "Source rows": len(it.source_records),
                "Invoice breakup": it.invoice_breakup,
                "Numeric total": it.aggregated_amount,
                "Excel formula": it.excel_formula_breakup,
                "Return": it.return_amount,
                "Matched row": it.matched_row,
                "Target cell": it.target_cell,
                "Existing value": it.existing_value,
                "Value to write": it.final_value,
                "Write mode": write_mode,
                "Status": it.status,
                "Messages": "; ".join(it.messages),
            })
        st.dataframe(pd.DataFrame(plan_rows), use_container_width=True)

        st.markdown("---")
        blocking = report.has_blocking_errors
        if blocking:
            st.error("Resolve blocking issues (unmapped groups / missing columns / errors) before writing.")
        confirm = st.checkbox("I have reviewed the plan above and confirm the write.")
        if st.button("✅ Write approved values to a NEW workbook", type="primary", disabled=blocking or not confirm):
            ops = plan_to_write_ops(report, CFG)
            out_bytes, outcomes = excel_writer.apply_writes(st.session_state["excel_bytes"], ops)
            st.session_state["output_bytes"] = out_bytes
            st.session_state["outcomes"] = outcomes
            st.session_state["audit_records"] = audit_mod.build_audit_records(
                outcomes, report.items, source_file=st.session_state["source_name"],
                timestamp=_dt.datetime.now().isoformat(timespec="seconds"),
            )
            written = sum(1 for o in outcomes if not o.skipped)
            st.success(f"Wrote {written} value(s) to a new workbook. Go to page 8 to download.")


# ===========================================================================
# PAGE 8 — EXPORT
# ===========================================================================
elif page == PAGES[7]:
    st.header("8. Export & download")
    if st.session_state["output_bytes"] is None:
        st.info("Nothing written yet. Complete page 7 and confirm the write.")
    else:
        st.success("Your updated workbook and audit files are ready.")
        st.download_button("⬇️ Updated Excel workbook",
                           data=st.session_state["output_bytes"],
                           file_name=f"updated_{st.session_state['excel_name'] or 'workbook.xlsx'}",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Backup of original.
        st.download_button("⬇️ Backup of original workbook",
                           data=excel_writer.make_backup(st.session_state["excel_bytes"]),
                           file_name=f"backup_{st.session_state['excel_name'] or 'workbook.xlsx'}",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        audit_records = st.session_state["audit_records"] or []
        st.download_button("⬇️ Audit report (CSV)",
                           data=audit_mod.audit_to_csv(audit_records),
                           file_name="audit.csv", mime="text/csv")
        st.download_button("⬇️ Audit report (JSON)",
                           data=audit_mod.audit_to_json(audit_records),
                           file_name="audit.json", mime="application/json")

        # Raw + aggregated exports.
        df = st.session_state["records_df"]
        if df is not None:
            st.download_button("⬇️ Extracted raw data (CSV)",
                               data=df.to_csv(index=False), file_name="raw_extracted.csv", mime="text/csv")
            records = df_to_records(df)
            rows = aggregate(records, CFG.source, aggregation_keys=CFG.write_rules.aggregation_keys)
            # Grouped totals now include numeric total + human breakup + excel formula breakup.
            agg_rows = grouped_export_rows(rows)
            st.download_button("⬇️ Grouped / aggregated data (CSV)",
                               data=audit_mod.records_to_csv(agg_rows),
                               file_name="grouped_totals.csv", mime="text/csv")

        # Write plan export (numeric total, breakup and formula, per target cell).
        report = st.session_state["plan"]
        if report is not None:
            st.download_button("⬇️ Write plan (CSV)",
                               data=audit_mod.records_to_csv(plan_to_export_rows(report, CFG)),
                               file_name="write_plan.csv", mime="text/csv")

        st.download_button("⬇️ Saved mapping (JSON)",
                           data=config_to_json(CFG), file_name="mapping.json", mime="application/json")

        # Error / warning report.
        if report is not None:
            warn_rows = [{"group": it.group, "status": it.status, "messages": "; ".join(it.messages)}
                         for it in report.items if it.status != Status.READY or it.messages]
            st.download_button("⬇️ Error / warning report (CSV)",
                               data=audit_mod.records_to_csv(warn_rows or [{"group": "", "status": "", "messages": ""}]),
                               file_name="warnings.csv", mime="text/csv")


st.session_state["config"] = CFG
