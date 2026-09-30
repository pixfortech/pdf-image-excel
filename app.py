"""PDF / image report -> update an existing Excel workbook.

Normal use is three steps: upload both files, review the preview, click
"Update uploaded workbook".  The mapping is recognised from the report's
structure and loaded from the local profile store; you are asked only about
what is new or has changed.

Run:  streamlit run app.py
"""
from __future__ import annotations

import copy
import json

import pandas as pd
import streamlit as st

from src import audit, pipeline, plan as plan_mod, profiles, values
from src.profiles import (ColumnRef, FORMULA, NUMERIC, OVERWRITE, SKIP_NONEMPTY,
                          Profile, ProfileStore, SheetLayout)

st.set_page_config(page_title="Update workbook from report", layout="wide")
ss = st.session_state
ss.setdefault("store_version", 0)
ss.setdefault("forced_profile", "")
ss.setdefault("result", None)


@st.cache_resource
def get_store() -> ProfileStore:
    return ProfileStore()


store = get_store()


def saved(profile: Profile, message: str) -> None:
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
    st.sidebar.caption(f"Stored locally in `{store.root}` (never uploaded to Git).")
    names = store.names()
    options = [""] + names
    current = ss.forced_profile if ss.forced_profile in names else ""
    choice = st.sidebar.selectbox(
        "Profile", options, index=options.index(current),
        format_func=lambda n: n or "Automatic — recognise from the report",
    )
    if choice != ss.forced_profile:
        ss.forced_profile, ss.result = choice, None
        st.rerun()
    last = store.last_successful()
    if last and st.sidebar.button("Use last successful mapping", use_container_width=True):
        ss.forced_profile, ss.result = last.profile_name, None
        st.rerun()

    uploaded = st.sidebar.file_uploader("Import profile JSON", type=["json"], key="import_json")
    if uploaded is not None and st.sidebar.button("Import", use_container_width=True):
        try:
            p = store.import_json(uploaded.getvalue().decode("utf-8"))
            ss.store_version += 1
            st.sidebar.success(f"Imported as '{p.profile_name}'.")
        except (ValueError, json.JSONDecodeError) as exc:
            st.sidebar.error(f"Not a valid profile: {exc}")

    if prepared is None or not store.load(prepared.profile.profile_name):
        return
    profile = store.load(prepared.profile.profile_name)
    st.sidebar.divider()
    st.sidebar.markdown(f"**Current:** {profile.profile_name}")
    st.sidebar.download_button("Export profile JSON", profiles.to_json(profile),
                               file_name=f"{profile.profile_name}.json", mime="application/json",
                               use_container_width=True)
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
        if st.checkbox("Reset this profile's customer and sheet mapping"):
            if st.button("Reset mapping"):
                profile.group_to_sheet, profile.ignored_groups = {}, []
                profile.sheet_template, profile.sheet_overrides = SheetLayout(), {}
                saved(profile, "Mapping reset.")
        if st.checkbox("Delete this profile permanently"):
            if st.button("Delete profile", type="primary"):
                store.delete(profile.profile_name)
                ss.forced_profile, ss.result = "", None
                ss.store_version += 1
                st.rerun()


# ===========================================================================
# Mapping editor (first run, exceptions, or Advanced)
# ===========================================================================

def _default_letter(headers: dict, current: ColumnRef):
    """Pre-select by the saved HEADER first (a template's 'RETURN' may be in a
    different column on another sheet); if that header is absent, pre-select
    nothing so the user must choose — never a same-letter column by accident."""
    if current.header:
        key = values.label_key(current.header)
        found = [l for l, h in headers.items() if values.label_key(h) == key]
        if len(found) == 1:
            return found[0]
        if current.letter and values.label_key(headers.get(current.letter, "")) == key:
            return current.letter
        return None
    return current.letter if current.letter in headers else None


def column_select(label, wb, sheet, header_row, current: ColumnRef, key, optional=False) -> ColumnRef:
    headers = dict(wb.headers(sheet, header_row))
    letters = list(headers)
    default = _default_letter(headers, current)
    index = letters.index(default) if default else None
    options = ([""] if optional else []) + letters
    if optional:
        index = 0 if index is None else index + 1
    choice = st.selectbox(label, options, index=index, key=key, placeholder="Choose a column",
                          format_func=lambda l: "— do not write —" if not l else f"{l} — {headers.get(l) or '(blank)'}")
    return ColumnRef(choice, headers.get(choice, "")) if choice else ColumnRef()


def layout_inputs(wb, sheet, layout: SheetLayout, key: str, returns_needed: bool) -> SheetLayout:
    c0, c1, c2, c3 = st.columns([1, 2, 2, 2])
    header_row = c0.number_input("Header row", 1, 50, int(layout.header_row or 1), key=f"{key}_hr")
    with c1:
        date = column_select("DATE column", wb, sheet, header_row, layout.date, f"{key}_date")
    with c2:
        amount = column_select("Main amount target", wb, sheet, header_row, layout.amount, f"{key}_amt")
    with c3:
        returns = column_select("Returns target", wb, sheet, header_row, layout.returns,
                                f"{key}_ret", optional=True) if returns_needed else layout.returns
    return SheetLayout(int(header_row), date, amount, returns)


def mapping_editor(prepared, full: bool) -> None:
    """One form. ``full`` = Advanced (everything); otherwise only what needs attention."""
    profile, report, wb = copy.deepcopy(prepared.profile), prepared.report, prepared.wb
    k = (lambda name: f"adv_{name}") if full else (lambda name: name)   # unique widget keys
    is_new = prepared.profile_source == pipeline.NEW
    sheet_names = wb.sheets if wb else []
    problem_groups = [s.group for s in prepared.plan.sheets
                      if s.status in (plan_mod.UNMAPPED, plan_mod.AMBIGUOUS, plan_mod.SHEET_MISSING)]
    problem_sheets = sorted({s.sheet for s in prepared.plan.sheets if s.status == plan_mod.COLUMN_PROBLEM})
    show_fields = full or is_new or bool(prepared.field_issues)
    show_template = full or is_new or not profile.sheet_template.amount.is_set
    groups_to_show = report.groups if (full or is_new) else problem_groups
    returns_needed = bool(profile.source.returns_field) or full

    with st.form(k("mapping")):
        if show_fields:
            st.markdown("**Report fields** — chosen by column label, validated by content")
            cols = report.columns
            samples = {c: next((r.values.get(c, "") for r in report.rows if r.values.get(c, "").strip()), "")
                       for c in cols}
            labels = {c: f"{c}  ·  {report.column_kind(c)}  ·  e.g. {samples[c] or '—'}" for c in cols}
            f1, f2, f3, f4 = st.columns(4)

            def pick(slot, title, current, optional):
                opts = ([""] if optional else []) + cols
                idx = opts.index(current) if current in opts else (0 if optional else None)
                return slot.selectbox(title, opts, index=idx, placeholder="Choose",
                                      format_func=lambda c: labels.get(c, "— none —"), key=k(f"fld_{title}"))
            date_f = pick(f1, "Date", profile.source.date_field, False)
            amount_f = pick(f2, "Main amount", profile.source.amount_field, False)
            returns_f = pick(f3, "Returns (optional)", profile.source.returns_field, True)
            ref_f = pick(f4, "Reference (audit only)", profile.source.reference_field, True)
            name = st.text_input("Profile name", value=profile.profile_name, key=k("name")) if is_new else profile.profile_name

        if groups_to_show:
            st.markdown("**Customer → worksheet** (remembered; names are matched ignoring case, "
                        "punctuation and spacing)")
            choices = {}
            grid = st.columns(3)
            opts = ["", "(ignore this customer)"] + sheet_names
            for i, group in enumerate(groups_to_show):
                match = profiles.lookup_group(profile, group)
                current = "(ignore this customer)" if match.status == "ignored" else match.sheet
                choices[group] = grid[i % 3].selectbox(
                    group, opts, index=opts.index(current) if current in opts else 0,
                    format_func=lambda s: s or "— choose worksheet —", key=k(f"grp_{values.name_key(group)}"))

        template_sheet, template = None, profile.sheet_template
        if show_template and wb:
            st.markdown("**Worksheet columns** — one template used by every mapped worksheet "
                        "whose headers match it")
            mapped = [s for s in profile.mapped_sheets() if wb.has_sheet(s)]
            sample_opts = mapped or sheet_names
            template_sheet = st.selectbox("Template worksheet", sample_opts, key=k("tmpl_sheet"))
            base = template if template.date.is_set else wb.suggest_layout(template_sheet)
            base.amount, base.returns = template.amount, template.returns
            template = layout_inputs(wb, template_sheet, base, k("tmpl"), returns_needed)

        overrides = {}
        override_sheets = sorted(set(problem_sheets) | (set(profile.sheet_overrides) if full else set()))
        if full:
            override_sheets = sorted(set(override_sheets) | set(profile.mapped_sheets()))
        override_sheets = [s for s in override_sheets if wb and wb.has_sheet(s)]
        if override_sheets:
            st.markdown("**Worksheet overrides** — for sheets whose columns differ from the template")
            for sheet in override_sheets:
                own = sheet in profile.sheet_overrides
                label = f"{sheet} — {'override' if own else 'uses template'}"
                if sheet in problem_sheets:
                    note = next(s.note for s in prepared.plan.sheets if s.sheet == sheet)
                    label = f"{sheet} — needs attention: {note}"
                with st.expander(label, expanded=sheet in problem_sheets):
                    use_own = st.checkbox("Use its own columns (override)", value=own or sheet in problem_sheets,
                                          key=k(f"ovr_on_{sheet}"))
                    layout = layout_inputs(wb, sheet, profile.layout_for(sheet), k(f"ovr_{sheet}"), returns_needed)
                    if use_own:
                        overrides[sheet] = layout

        submitted = st.form_submit_button("Save mapping", type="primary")

    if not submitted:
        return
    if show_fields:
        if not date_f or not amount_f:
            st.error("Choose the date and main amount fields.")
            return
        profile.profile_name = (name or "New profile").strip()
        profile.source.date_field, profile.source.amount_field = date_f, amount_f
        profile.source.returns_field, profile.source.reference_field = returns_f or "", ref_f or ""
        profile.source.columns = list(report.columns)
        profile.source.group_label = report.group_label
        profile.source.group_total_label = report.group_total_label
    if groups_to_show:
        for group, sheet in choices.items():
            if sheet == "(ignore this customer)":
                profiles.ignore_group(profile, group)
            elif sheet:
                profiles.remember_group(profile, group, sheet)
    if show_template and template_sheet:
        if not template.date.is_set or not template.amount.is_set:
            st.error("Choose the DATE column and the main amount target column.")
            return
        profile.sheet_template = template
    for sheet in override_sheets:
        if sheet in overrides:
            profile.sheet_overrides[sheet] = overrides[sheet]
        else:
            profile.sheet_overrides.pop(sheet, None)
    saved(profile, "Mapping saved.")


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
            saved(profile, "Settings saved.")


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
    st.warning("No saved mapping matches this report yet. Confirm it once below — it is then remembered.")

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
    st.error(problem)

needs_attention = (prepared.profile_source == pipeline.NEW or prepared.field_issues
                   or any(s.status in (plan_mod.UNMAPPED, plan_mod.AMBIGUOUS, plan_mod.SHEET_MISSING,
                                       plan_mod.COLUMN_PROBLEM) for s in prepared.plan.sheets))
if needs_attention and prepared.wb:
    st.markdown("#### Needs attention")
    mapping_editor(prepared, full=False)

if prepared.plan.sheets:
    st.markdown("#### Mapping")
    st.dataframe(pd.DataFrame([{
        rep.group_label or "Group": s.group, "Worksheet": s.sheet,
        "Date matches": f"{s.dates_matched}/{s.dates_total}", "Target": s.target,
        "Status": s.status, "Note": s.note,
    } for s in prepared.plan.sheets]), use_container_width=True, hide_index=True)

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
    st.dataframe(table, use_container_width=True, hide_index=True, height=420)

with st.expander("Advanced settings and diagnostics"):
    tabs = st.tabs(["Mapping", "Write settings", "Reconciliation", "Rejected lines", "Exports"])
    with tabs[0]:
        if prepared.wb:
            mapping_editor(prepared, full=True)
    with tabs[1]:
        settings_editor(copy.deepcopy(prepared.profile))
    with tabs[2]:
        st.dataframe(pd.DataFrame(audit.reconciliation_rows(recon)), use_container_width=True, hide_index=True)
    with tabs[3]:
        st.caption("Lines that are not invoices (titles, headers, page footers, printed totals).")
        st.dataframe(pd.DataFrame([{"Page": r.page, "Line": r.text, "Reason": r.reason}
                                   for r in rep.ignored + prepared.rejected]),
                     use_container_width=True, hide_index=True)
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
        st.warning(f"{len(blocked)} cells cannot be written (see statuses above).")
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
