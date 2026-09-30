"""PDF / image report -> update an existing Excel workbook.

Normal use: upload the report and the workbook, look at the preview, click
"Update uploaded workbook".  The saved profile (customer -> worksheet and one
common worksheet pattern) is recognised and applied automatically; only
exceptions are shown.

Run:  streamlit run app.py
"""
from __future__ import annotations

import copy
import json
from typing import Optional

import pandas as pd
import streamlit as st

from src import audit, pipeline, plan as plan_mod, profiles, values
from src.layout import resolve_layout
from src.profiles import (ColumnRef, FORMULA, NUMERIC, OVERWRITE, SKIP_NONEMPTY,
                          Profile, ProfileStore, SheetLayout)

st.set_page_config(page_title="Update workbook from report", layout="wide")
ss = st.session_state
ss.setdefault("store_version", 0)
ss.setdefault("forced_profile", "")
ss.setdefault("result", None)

IGNORE = "(ignore this customer)"
GROUP_PROBLEMS = (plan_mod.UNMAPPED, plan_mod.AMBIGUOUS, plan_mod.SHEET_MISSING)


@st.cache_resource
def get_store() -> ProfileStore:
    return ProfileStore()


store = get_store()


def save(profile: Profile, message: str) -> None:
    store.save(profile)
    ss.store_version += 1
    ss.result = None
    st.toast(message)
    st.rerun()


# ===========================================================================
# Sidebar: profiles
# ===========================================================================

def sidebar(prepared) -> None:
    st.sidebar.header("Mapping profiles")
    st.sidebar.caption(f"Saved on this computer in `{store.root}` (never uploaded to Git).")
    names = store.names()
    options = [""] + names
    current = ss.forced_profile if ss.forced_profile in names else ""
    choice = st.sidebar.selectbox("Profile", options, index=options.index(current),
                                  format_func=lambda n: n or "Automatic — recognise from the report")
    if choice != ss.forced_profile:
        ss.forced_profile, ss.result = choice, None
        st.rerun()
    last = store.last_successful()
    if last and st.sidebar.button("Use last successful mapping", width="stretch"):
        ss.forced_profile, ss.result = last.profile_name, None
        st.rerun()

    uploaded = st.sidebar.file_uploader("Import profile JSON (backup / other computer)", type=["json"])
    if uploaded is not None and st.sidebar.button("Import", width="stretch"):
        try:
            p = store.import_json(uploaded.getvalue().decode("utf-8"))
            ss.store_version += 1
            st.sidebar.success(f"Imported as '{p.profile_name}'.")
        except (ValueError, json.JSONDecodeError) as exc:
            st.sidebar.error(f"Not a valid profile: {exc}")

    profile = store.load(prepared.profile.profile_name) if prepared else None
    if profile is None:
        return
    st.sidebar.divider()
    st.sidebar.markdown(f"**Current:** {profile.profile_name}")
    st.sidebar.download_button("Export profile JSON", profiles.to_json(profile),
                               file_name=f"{profile.profile_name}.json", mime="application/json",
                               width="stretch")
    with st.sidebar.expander("Rename / reset / delete"):
        new_name = st.text_input("New name", value=profile.profile_name)
        if st.button("Rename") and new_name.strip() and new_name != profile.profile_name:
            try:
                store.rename(profile.profile_name, new_name.strip())
                ss.forced_profile = new_name.strip() if ss.forced_profile else ""
                ss.store_version += 1
                st.rerun()
            except ValueError as exc:
                st.error(str(exc))
        if st.checkbox("Reset this profile's customer and worksheet mapping"):
            if st.button("Reset mapping"):
                profile.group_to_sheet, profile.ignored_groups = {}, []
                profile.sheet_template, profile.sheet_overrides, profile.template_sheet = SheetLayout(), {}, ""
                save(profile, "Mapping reset.")
        if st.checkbox("Delete this profile permanently"):
            if st.button("Delete profile", type="primary"):
                store.delete(profile.profile_name)
                ss.forced_profile, ss.result = "", None
                ss.store_version += 1
                st.rerun()


# ===========================================================================
# Column pickers — always built from the CURRENT worksheet's real headers
# ===========================================================================

def _default_letter(headers: dict, ref: ColumnRef) -> Optional[str]:
    """Pre-select by saved header name first (a pattern's RETURN may sit in a
    different column on another sheet); if absent, pre-select nothing."""
    if ref.header:
        key = values.label_key(ref.header)
        found = [l for l, h in headers.items() if h and values.label_key(h) == key]
        if len(found) == 1:
            return found[0]
        if ref.letter and values.label_key(headers.get(ref.letter, "")) == key:
            return ref.letter
        return None
    return ref.letter if headers.get(ref.letter) else None


def column_picker(label, wb, sheet, header_row, ref: ColumnRef, key_prefix, field, optional=False) -> ColumnRef:
    """The widget key contains the worksheet and header row, so choosing a
    different worksheet always shows that worksheet's columns — never stale ones."""
    headers = {l: h for l, h in wb.headers(sheet, header_row) if h.strip()}
    letters = list(headers)
    default = _default_letter(headers, ref)
    options = ([""] if optional else []) + letters
    index = options.index(default) if default in options else (0 if optional else None)
    choice = st.selectbox(label, options, index=index, placeholder="Choose a column",
                          key=f"{key_prefix}_{sheet}_{header_row}_{field}",
                          format_func=lambda l: "— do not write —" if not l else f"{l} — {headers[l]}")
    return ColumnRef(choice, headers[choice]) if choice else ColumnRef()


def layout_editor(wb, sheet, base: SheetLayout, key_prefix, amount_label, returns_label) -> SheetLayout:
    c0, c1, c2, c3 = st.columns([1, 2, 2, 2])
    header_row = int(c0.number_input("Header row", 1, 50, int(base.header_row or 1),
                                     key=f"{key_prefix}_{sheet}_hr"))
    with c1:
        date = column_picker("DATE column (row matching)", wb, sheet, header_row, base.date, key_prefix, "date")
    with c2:
        amount = column_picker(f"PDF '{amount_label}' goes to", wb, sheet, header_row, base.amount,
                               key_prefix, "amount")
    with c3:
        returns = (column_picker(f"PDF '{returns_label}' goes to", wb, sheet, header_row, base.returns,
                                 key_prefix, "returns", optional=True)
                   if returns_label else ColumnRef())
    return SheetLayout(header_row, date, amount, returns)


def pattern_lines(profile: Profile, prepared) -> list:
    """'DATE -> A / DATE' lines for the common pattern, as found on its template sheet."""
    t, src = profile.sheet_template, profile.source
    res = None
    for sheet in [profile.template_sheet] + profile.mapped_sheets():
        if prepared.wb.has_sheet(sheet):
            res = resolve_layout(prepared.wb, sheet, t, bool(t.returns.is_set))
            if res.ok:
                break

    def show(ref: ColumnRef, letter):
        return f"{letter} / {ref.header}" if letter else ref.describe()
    lines = [f"DATE → {show(t.date, res and res.date)}",
             f"{src.amount_field} → {show(t.amount, res and res.amount)}"]
    if src.returns_field:
        lines.append(f"{src.returns_field} → "
                     + (show(t.returns, res and res.returns) if t.returns.is_set else "not written"))
    return lines


# ===========================================================================
# Mapping (normal view: summary + exceptions only)
# ===========================================================================

def customer_editor(prepared, groups, key_prefix, title) -> None:
    profile, sheets = prepared.profile, prepared.wb.sheets
    with st.form(f"{key_prefix}_customers"):
        st.markdown(title)
        grid = st.columns(3)
        options = ["", IGNORE] + sheets
        choices = {}
        for i, group in enumerate(groups):
            match = profiles.lookup_group(profile, group)
            current = IGNORE if match.status == "ignored" else match.sheet
            choices[group] = grid[i % 3].selectbox(
                group, options, index=options.index(current) if current in options else 0,
                format_func=lambda s: s or "— choose worksheet —",
                key=f"{key_prefix}_cust_{values.name_key(group)}")
        if st.form_submit_button("Save customer mapping", type="primary"):
            updated = copy.deepcopy(profile)
            for group, sheet in choices.items():
                if sheet == IGNORE:
                    profiles.ignore_group(updated, group)
                elif sheet:
                    profiles.remember_group(updated, group, sheet)
            save(updated, "Customer mapping saved.")


def fields_editor(prepared, key_prefix) -> None:
    """Report fields: which printed column is the date / amount / returns."""
    report, profile = prepared.report, prepared.profile
    cols = report.columns
    sample = {c: next((r.values.get(c, "") for r in report.rows if r.values.get(c, "").strip()), "") for c in cols}
    label = {c: f"{c} · {report.column_kind(c)} · e.g. {sample[c] or '—'}" for c in cols}
    with st.form(f"{key_prefix}_fields"):
        f1, f2, f3, f4 = st.columns(4)

        def pick(slot, title, current, optional):
            opts = ([""] if optional else []) + cols
            idx = opts.index(current) if current in opts else (0 if optional else None)
            return slot.selectbox(title, opts, index=idx, placeholder="Choose",
                                  format_func=lambda c: label.get(c, "— none —"), key=f"{key_prefix}_fld_{title}")
        date_f = pick(f1, "Date", profile.source.date_field, False)
        amount_f = pick(f2, "Main amount", profile.source.amount_field, False)
        returns_f = pick(f3, "Returns (optional)", profile.source.returns_field, True)
        ref_f = pick(f4, "Reference (audit only)", profile.source.reference_field, True)
        if st.form_submit_button("Save report fields"):
            if not date_f or not amount_f:
                st.error("Choose the date and main amount fields.")
                return
            updated = copy.deepcopy(profile)
            s = updated.source
            s.date_field, s.amount_field, s.returns_field, s.reference_field = date_f, amount_f, returns_f or "", ref_f or ""
            s.columns, s.group_label, s.group_total_label = list(cols), report.group_label, report.group_total_label
            save(updated, "Report fields saved.")


def pattern_setup(prepared, key_prefix, allow_all_sheets=False) -> None:
    """Choose the common worksheet pattern.  NOT inside a form: changing the
    worksheet immediately reloads that worksheet's headers."""
    profile, wb = prepared.profile, prepared.wb
    mapped = [s for s in profile.mapped_sheets() if wb.has_sheet(s)]
    candidates = mapped + ([s for s in wb.sheets if s not in mapped] if allow_all_sheets else [])
    if not candidates:
        st.info("Map customers to worksheets first; the worksheet pattern is taken from a mapped worksheet.")
        return
    suggestion = prepared.pattern
    preferred = (profile.template_sheet if profile.template_sheet in candidates
                 else suggestion.template_sheet if suggestion and suggestion.template_sheet in candidates
                 else candidates[0])
    sheet = st.selectbox("Take the pattern from worksheet", candidates, index=candidates.index(preferred),
                         key=f"{key_prefix}_pattern_sheet")
    if profile.sheet_template.amount.is_set:
        base = copy.deepcopy(profile.sheet_template)
    elif suggestion and sheet == suggestion.template_sheet:
        base = copy.deepcopy(suggestion.layout)
    else:
        base = wb.suggest_layout(sheet)
    layout = layout_editor(wb, sheet, base, f"{key_prefix}_pattern",
                           profile.source.amount_field, profile.source.returns_field)
    if st.button("Save worksheet pattern", type="primary", key=f"{key_prefix}_pattern_save"):
        if not layout.date.is_set or not layout.amount.is_set:
            st.error("Choose the DATE column and the column the main amount goes to.")
            return
        updated = copy.deepcopy(profile)
        updated.sheet_template, updated.template_sheet = layout, sheet
        save(updated, f"Worksheet pattern saved from {sheet}.")


def exception_editor(prepared, sheet, key_prefix, note="") -> None:
    """Own columns for one worksheet whose layout differs from the pattern."""
    profile, wb = prepared.profile, prepared.wb
    if note:
        st.caption(f"Why: {note}")
    base = copy.deepcopy(profile.layout_for(sheet))
    layout = layout_editor(wb, sheet, base, f"{key_prefix}_exc",
                           profile.source.amount_field, profile.source.returns_field)
    c1, c2 = st.columns(2)
    if c1.button(f"Save for {sheet}", type="primary", key=f"{key_prefix}_exc_{sheet}_save"):
        if not layout.date.is_set or not layout.amount.is_set:
            st.error("Choose the DATE column and the column the main amount goes to.")
            return
        updated = copy.deepcopy(profile)
        updated.sheet_overrides[sheet] = layout
        save(updated, f"{sheet} saved as an exception.")
    if sheet in profile.sheet_overrides and c2.button(f"Use the common pattern for {sheet}",
                                                      key=f"{key_prefix}_exc_{sheet}_drop"):
        updated = copy.deepcopy(profile)
        updated.sheet_overrides.pop(sheet, None)
        save(updated, f"{sheet} now follows the common pattern.")


def mapping_section(prepared) -> None:
    profile, plan = prepared.profile, prepared.plan
    st.markdown("#### Mapping")

    if prepared.field_issues:
        st.error("The report's fields need attention: " + " ".join(prepared.field_issues.values()))
        fields_editor(prepared, "fix")
        return

    # Customer -> Worksheet
    if plan.sheets:
        st.markdown("**Customer → Worksheet**")
        st.dataframe(pd.DataFrame([{
            prepared.report.group_label or "Customer": s.group,
            "Worksheet": s.sheet or "—",
            "Dates found": f"{s.dates_matched}/{s.dates_total}" if s.sheet else "",
            "Writes to": s.target or "—",
            "Status": s.status + (" (exception)" if s.via == "exception" else ""),
        } for s in plan.sheets]), width="stretch", hide_index=True)
    unmapped = [s.group for s in plan.sheets if s.status in GROUP_PROBLEMS]
    if unmapped:
        everything = len(unmapped) == len(plan.sheets)
        customer_editor(prepared, prepared.report.groups if everything else unmapped, "fix",
                        "**Map customers to worksheets** (saved; names match ignoring case, "
                        "punctuation and spacing)" if everything else
                        f"**Needs attention — {len(unmapped)} customer(s) without a worksheet**")

    # Common worksheet pattern
    st.markdown("**Worksheet pattern**")
    if profile.sheet_template.amount.is_set:
        mapped = [s for s in profile.mapped_sheets() if prepared.wb.has_sheet(s)]
        applied = {s.sheet for s in plan.sheets if s.via == "pattern" and s.status != plan_mod.COLUMN_PROBLEM}
        st.markdown("  \n".join(pattern_lines(profile, prepared)))
        st.caption(f"Applied automatically to "
                   f"{len(applied)} of {len(mapped)} mapped worksheets"
                   + (f" · {len(profile.sheet_overrides)} exception(s)" if profile.sheet_overrides else ""))
    else:
        s = prepared.pattern
        if s:
            found = [f"DATE → {s.layout.date.letter} / {s.layout.date.header}"]
            if s.layout.returns.is_set:
                found.append(f"{profile.source.returns_field} → {s.layout.returns.letter} / {s.layout.returns.header}")
            st.info("Detected on worksheet **" + s.template_sheet + "**: " + " · ".join(found)
                    + f". Choose once where the PDF's **{profile.source.amount_field}** goes; "
                      "the pattern is then applied to every mapped worksheet with the same headers.")
        pattern_setup(prepared, "fix")

    # Exceptions: only the worksheets whose headers differ
    if profile.sheet_template.amount.is_set:
        problems = {}
        for s in plan.sheets:
            if s.status == plan_mod.COLUMN_PROBLEM and s.sheet:
                problems.setdefault(s.sheet, s.note)
        if problems:
            st.markdown(f"**Needs attention — {len(problems)} worksheet(s) laid out differently**")
            for sheet, note in problems.items():
                with st.expander(f"{sheet}: {note}", expanded=len(problems) <= 3):
                    exception_editor(prepared, sheet, "fix")


def advanced_mapping(prepared) -> None:
    profile, wb = prepared.profile, prepared.wb
    st.markdown("**Report fields**")
    fields_editor(prepared, "adv")
    customer_editor(prepared, prepared.report.groups, "adv", "**All customers → worksheet**")
    st.markdown("**Common worksheet pattern**")
    pattern_setup(prepared, "adv", allow_all_sheets=True)
    st.markdown("**Worksheet exceptions**")
    mapped = [s for s in profile.mapped_sheets() if wb.has_sheet(s)]
    if mapped:
        sheet = st.selectbox("Worksheet", mapped, key="adv_exc_sheet",
                             format_func=lambda s: f"{s} — {'exception' if s in profile.sheet_overrides else 'common pattern'}")
        exception_editor(prepared, sheet, "adv")


def settings_editor(profile: Profile) -> None:
    with st.form("settings"):
        w = profile.write
        c1, c2 = st.columns(2)
        output = c1.radio("Write", [NUMERIC, FORMULA], index=[NUMERIC, FORMULA].index(w.output_mode),
                          format_func=lambda v: {NUMERIC: "Numeric total (19,750)",
                                                 FORMULA: "Formula breakup (=12500+7250)"}[v])
        existing = c2.radio("If a target cell already has a different value",
                            [OVERWRITE, SKIP_NONEMPTY], index=[OVERWRITE, SKIP_NONEMPTY].index(w.existing_values),
                            format_func=lambda v: {OVERWRITE: "Overwrite (shown in preview)",
                                                   SKIP_NONEMPTY: "Keep the existing value"}[v])
        c3, c4 = st.columns(2)
        write_returns = c3.checkbox("Write Returns when the report has a Returns value", value=w.write_returns)
        zero_fill = c4.checkbox("Write 0 into Returns when the report's Returns is blank",
                                value=w.zero_fill_returns,
                                help="Off by default: returns may come from a separate Sales Return report.")
        order = st.radio("Ambiguous dates like 05/06/2026", list(values.DATE_ORDERS),
                         index=list(values.DATE_ORDERS).index(profile.source.date_order), horizontal=True,
                         format_func=lambda v: {values.DMY: "Day first (India/UK)", values.MDY: "Month first (US)",
                                                values.AUTO: "Auto-detect"}[v])
        if st.form_submit_button("Save settings"):
            profile.write.output_mode, profile.write.existing_values = output, existing
            profile.write.write_returns, profile.write.zero_fill_returns = write_returns, zero_fill
            profile.source.date_order = order
            save(profile, "Settings saved.")


# ===========================================================================
# Stage 1 — Upload
# ===========================================================================

st.title("Update workbook from report")
st.subheader("1 · Upload")
c1, c2 = st.columns(2)
pdf_file = c1.file_uploader("Report (PDF or image)",
                            type=["pdf", "png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff"])
xlsx_file = c2.file_uploader("Existing Excel workbook to update", type=["xlsx", "xlsm"])

if not (pdf_file and xlsx_file):
    sidebar(None)
    st.info("Upload the report and the workbook. The saved mapping is found automatically.")
    st.stop()

pdf_bytes, xlsx_bytes = pdf_file.getvalue(), xlsx_file.getvalue()
key = (pipeline.digest(pdf_bytes), pipeline.digest(xlsx_bytes), ss.forced_profile, ss.store_version)
if ss.get("prep_key") != key:
    with st.spinner("Reading the report and the workbook…"):
        forced = store.load(ss.forced_profile) if ss.forced_profile else None
        ss.prepared = pipeline.prepare(pdf_bytes, pdf_file.name, xlsx_bytes, xlsx_file.name, store, forced)
    ss.prep_key, ss.result = key, None
prepared: pipeline.Prepared = ss.prepared
sidebar(prepared)

# ===========================================================================
# Stage 2 — Review
# ===========================================================================

st.subheader("2 · Review")
if prepared.profile_source == pipeline.LOADED:
    st.success(f"Previous mapping recognised and loaded: **{prepared.profile.profile_name}**")
elif prepared.profile_source == pipeline.GIVEN:
    st.info(f"Using profile **{prepared.profile.profile_name}**.")
else:
    st.warning("No saved mapping for this report yet. Set it up once below — it is then remembered.")

rep, recon = prepared.report, prepared.reconciliation
period = f"{values.format_date(rep.period[0])} – {values.format_date(rep.period[1])}" \
    if rep.period and rep.period[1] else "—"
m = st.columns(5)
m[0].metric("Report period", period)
m[1].metric(rep.group_label or "Groups", len(rep.groups))
m[2].metric("Invoices", len(prepared.invoices))
m[3].metric("Printed totals", "Reconciled" if recon and recon.checked and recon.ok
            else ("Mismatch" if recon and not recon.ok else "Not checked"))
m[4].metric("Cells ready", prepared.ready_count)

for warning in prepared.document.warnings + rep.warnings + (recon.warnings if recon else []):
    st.warning(warning)
for problem in prepared.hard_blockers:
    if not prepared.field_issues or problem not in prepared.field_issues.values():
        st.error(problem)

if prepared.wb:
    mapping_section(prepared)

if prepared.plan.rows:
    st.markdown("#### What will be written")
    show_all = st.toggle("Show unchanged and ignored rows too", value=False)
    formula_mode = prepared.profile.write.output_mode == FORMULA
    rows = [r for r in prepared.plan.rows
            if show_all or r.is_ready or r.is_blocking or r.status == plan_mod.KEEP_EXISTING]
    table = pd.DataFrame([{
        rep.group_label or "Group": x["customer"], "Date": x["date"], "Field": x["field"],
        "Invoices": x["invoice_count"], "Breakup": x["breakup"], "Aggregate": x["aggregate"],
        **({"Formula": x["formula"]} if formula_mode else {}),
        "Worksheet": x["worksheet"], "Cell": x["cell"], "Existing": x["existing_value"],
        "New value": x["new_value"], "Status": x["status"],
    } for x in audit.plan_rows(rows)])
    st.dataframe(table, width="stretch", hide_index=True, height=420)

with st.expander("Advanced mapping, settings and diagnostics"):
    tabs = st.tabs(["Advanced mapping", "Write settings", "Reconciliation", "Rejected lines", "Exports"])
    with tabs[0]:
        if prepared.wb:
            advanced_mapping(prepared)
    with tabs[1]:
        settings_editor(copy.deepcopy(prepared.profile))
    with tabs[2]:
        st.dataframe(pd.DataFrame(audit.reconciliation_rows(recon)), width="stretch", hide_index=True)
    with tabs[3]:
        st.caption("Lines that are not invoices (titles, headers, page footers, printed totals).")
        st.dataframe(pd.DataFrame([{"Page": r.page, "Line": r.text, "Reason": r.reason}
                                   for r in rep.ignored + prepared.rejected]),
                     width="stretch", hide_index=True)
        st.caption(f"Extraction: {prepared.document.engine}. Columns: " + ", ".join(
            f"{c} ({rep.column_kind(c)})" for c in rep.columns))
    with tabs[4]:
        for label, rows_, name in (
            ("Extracted invoices", audit.invoice_rows(prepared.invoices), "invoices.csv"),
            ("Write plan", audit.plan_rows(prepared.plan.rows), "write_plan.csv"),
            ("Reconciliation", audit.reconciliation_rows(recon), "reconciliation.csv"),
        ):
            if rows_:
                st.download_button(label, audit.to_csv(rows_), file_name=name, mime="text/csv")

# ===========================================================================
# Stage 3 — Update
# ===========================================================================

st.subheader("3 · Update")
ready = prepared.plan.ready
blocked = prepared.plan.blocking
overwrites = sum(r.status == plan_mod.READY_OVERWRITE for r in ready)
if prepared.hard_blockers:
    st.error("Resolve the problems above before updating the workbook.")
elif not ready:
    unchanged = sum(r.status == plan_mod.UNCHANGED for r in prepared.plan.rows)
    st.error("Nothing is ready to write." + (f" {unchanged} cells already hold these values." if unchanged else ""))
else:
    exclude = False
    if blocked:
        st.warning(f"{len(blocked)} cells cannot be written yet (see Needs attention above).")
        exclude = st.checkbox(f"Leave those {len(blocked)} cells out and update the other {len(ready)}")
    sheets = sorted({r.sheet for r in ready})
    confirm = st.checkbox(
        f"I reviewed the preview: {len(ready) - overwrites} empty cells will be filled"
        + (f" and **{overwrites} existing values overwritten**" if overwrites else "")
        + f" in {len(sheets)} worksheets.")
    if st.button("Update uploaded workbook", type="primary",
                 disabled=not (confirm and prepared.can_update(exclude))):
        try:
            ss.result = pipeline.update(prepared, store, exclude_blocked=exclude)
        except pipeline.UpdateError as exc:
            st.error(str(exc))

result = ss.result
if result is not None:
    touched = sorted({c.sheet for c in result.changes})
    st.success(f"Updated {len(result.changes)} cells in {len(touched)} worksheets "
               f"({', '.join(touched)}). Verified: every other part of the workbook is unchanged; "
               "Excel recalculates dependent formulas when the file is opened.")
    rows = audit.audit_rows(prepared, result)
    d1, d2, d3, d4 = st.columns(4)
    d1.download_button(f"Download Updated_{prepared.xlsx_name}", result.workbook,
                       file_name=f"Updated_{prepared.xlsx_name}", type="primary",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d2.download_button("Audit (CSV)", audit.to_csv(rows), file_name="audit.csv", mime="text/csv")
    d3.download_button("Audit (JSON)", audit.to_json(rows), file_name="audit.json", mime="application/json")
    d4.download_button("Original workbook (backup)", prepared.original, file_name=prepared.xlsx_name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
