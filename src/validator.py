"""Validation and build of the final write plan.

Combines aggregated source rows with the Excel mapping to produce a list of
:class:`PlanItem` rows, each annotated with a status (Ready, Unmapped Group,
Missing Column, Date Not Found, Conflict, Error).  It also resolves the actual
:class:`~src.excel_writer.WriteOp` objects for the items that are ready.

No worksheet, column, customer or value is hardcoded; everything comes from the
:class:`~src.mapping.AppConfig` provided by the user.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional

from . import excel_reader
from .aggregator import AggregatedRow, build_formula
from .excel_writer import WriteOp
from .mapping import (
    AppConfig,
    DateNotFoundAction,
    OutputType,
    TargetMode,
    WriteAction,
)
from . import utils


class Status:
    READY = "Ready"
    UNMAPPED_GROUP = "Unmapped Group"
    MISSING_COLUMN = "Missing Column"
    SHEET_NOT_FOUND = "Sheet Not Found"
    DATE_NOT_FOUND = "Date Not Found"
    CONFLICT = "Conflict"
    ERROR = "Error"
    INVALID_ROW = "Skipped (no valid date/amount)"

# Fraction of PDF dates that may be missing before a strong warning is shown.
DATE_MISS_WARN_THRESHOLD = 0.25


@dataclass
class PlanItem:
    group: str
    sheet: str = ""
    date: object = None
    date_raw: str = ""
    matched_row: Optional[int] = None
    target_column: Optional[int] = None
    target_column_letter: str = ""
    target_cell: str = ""
    source_amounts: List[float] = field(default_factory=list)
    aggregated_amount: float = 0.0          # numeric total, ALWAYS kept
    invoice_breakup: str = ""               # human-readable, e.g. "630.00 + 12,705.00"
    excel_formula_breakup: str = ""         # formula, e.g. "=630+12705"
    return_amount: float = 0.0
    existing_value: object = None
    final_value: object = None              # numeric total or formula (per write mode)
    write_action: str = ""
    status: str = Status.READY
    messages: List[str] = field(default_factory=list)
    return_column: Optional[int] = None
    return_existing_value: object = None
    source_records: List[dict] = field(default_factory=list)
    is_insert: bool = False
    is_append: bool = False          # appended below existing data (no physical row insert)
    template_row: Optional[int] = None
    date_column: Optional[int] = None  # resolved date column index (for inserts)
    # Date transparency (for the final preview):
    pdf_date_normalised: str = ""      # e.g. "2026-05-15"
    excel_date_value: object = None    # original Excel cell value at the matched row
    excel_date_normalised: str = ""    # normalised Excel date


@dataclass
class ValidationReport:
    items: List[PlanItem] = field(default_factory=list)
    unmapped_groups: List[str] = field(default_factory=list)
    dates_not_found: int = 0
    missing_columns: int = 0
    conflicts: int = 0
    errors: int = 0
    skipped_invalid: int = 0          # rows lacking a valid date AND amount
    duplicate_headers: Dict[str, List[str]] = field(default_factory=dict)
    date_match_warnings: List[str] = field(default_factory=list)

    @property
    def has_blocking_errors(self) -> bool:
        return (
            bool(self.unmapped_groups)
            or self.missing_columns > 0
            or self.errors > 0
        )

    @property
    def ready_count(self) -> int:
        return sum(1 for i in self.items if i.status == Status.READY)


class MapStatus:
    READY = "Ready"
    NEEDS_REVIEW = "Needs review"
    SHEET_MISSING = "Sheet missing"
    DATE_MISSING = "Date column missing"
    AMOUNT_MISSING = "Amount column missing"
    RETURN_MISSING = "Return column missing"


def sheet_mapping_status(wb, config: AppConfig) -> List[dict]:
    """Per-Customer-Name view of the Excel column mapping and its status.

    Returns one row per detected group->sheet assignment with columns:
    customer_name, assigned_sheet, header_row, date_column, amount_column,
    return_column, status.  Resolution honours header-name / column-letter /
    cell-reference modes, so repeated headers are handled.
    """
    sheet_names = set(excel_reader.list_sheets(wb))
    out: List[dict] = []
    for group, sheet in config.group_to_sheet.items():
        row = {
            "customer_name": group,
            "assigned_sheet": sheet,
            "header_row": "",
            "date_column": "",
            "amount_column": "",
            "return_column": "",
            "status": MapStatus.NEEDS_REVIEW,
        }
        if not sheet:
            out.append(row)
            continue
        if sheet not in sheet_names:
            row["status"] = MapStatus.SHEET_MISSING
            out.append(row)
            continue
        sm = config.sheets.get(sheet)
        if sm is None:
            row["status"] = MapStatus.NEEDS_REVIEW
            out.append(row)
            continue

        row["header_row"] = sm.header_row
        row["date_column"] = sm.date_column
        row["amount_column"] = sm.amount_cell or sm.amount_column
        row["return_column"] = sm.return_cell or sm.return_column

        # Resolve amount.
        if sm.amount_target_mode == TargetMode.CELL_REFERENCE and sm.amount_cell:
            amount_ok = excel_reader.cell_reference_to_rc(sm.amount_cell) is not None
        else:
            amount_ok = excel_reader.resolve_column_index(
                wb, sheet, mode=sm.amount_target_mode.value,
                selector=sm.amount_column, header_row=sm.header_row) is not None
        date_ok = excel_reader.resolve_column_index(
            wb, sheet, mode=sm.date_target_mode.value,
            selector=sm.date_column, header_row=sm.header_row) is not None or not sm.date_column

        return_ok = True
        if config.source.return_field and (sm.return_column or sm.return_cell):
            if sm.return_target_mode == TargetMode.CELL_REFERENCE and sm.return_cell:
                return_ok = excel_reader.cell_reference_to_rc(sm.return_cell) is not None
            else:
                return_ok = excel_reader.resolve_column_index(
                    wb, sheet, mode=sm.return_target_mode.value,
                    selector=sm.return_column, header_row=sm.header_row) is not None

        if not amount_ok:
            row["status"] = MapStatus.AMOUNT_MISSING
        elif not date_ok:
            row["status"] = MapStatus.DATE_MISSING
        elif not return_ok:
            row["status"] = MapStatus.RETURN_MISSING
        else:
            row["status"] = MapStatus.READY
        out.append(row)
    return out


def configured_mapped_sheets(config: AppConfig) -> List[str]:
    """Mapped worksheets (assigned to a group) that have a column mapping."""
    mapped = {s for s in config.group_to_sheet.values() if s}
    return [s for s in mapped if s in config.sheets]


def build_plan(
    wb,
    rows: List[AggregatedRow],
    config: AppConfig,
) -> ValidationReport:
    report = ValidationReport()
    sheet_names = set(excel_reader.list_sheets(wb))

    # Pre-compute duplicate header warnings per mapped sheet.
    for sheet_name, sm in config.sheets.items():
        if sheet_name in sheet_names:
            dups = excel_reader.find_duplicate_headers(wb, sheet_name, sm.header_row)
            if dups:
                report.duplicate_headers[sheet_name] = dups

    # Per-sheet cursor so that multiple missing dates appended to the same sheet
    # land on distinct, sequential rows instead of all colliding at max_row+1.
    append_cursor: Dict[str, int] = {}
    interp = config.source.date_interpretation or "dmy"
    # Track, per sheet, the set of PDF dates and how many matched (for the
    # >25%-missing warning and the date-column suggestions).
    per_sheet_targets: Dict[str, set] = {}
    per_sheet_missing: Dict[str, int] = {}
    # Cache the full date-column scan per (sheet, col): (date_set, min, max).
    date_col_cache: Dict[tuple, tuple] = {}

    def _date_col_info(sheet, col, sm):
        key = (sheet, col)
        if key not in date_col_cache:
            dates = excel_reader.column_dates(
                wb, sheet, col, header_row=sm.header_row,
                date_formats=sm.date_formats or None, interpretation=interp)
            parsed = [d for _, _, d in dates]
            date_col_cache[key] = (set(parsed),
                                   min(parsed) if parsed else None,
                                   max(parsed) if parsed else None)
        return date_col_cache[key]

    for agg in rows:
        item = PlanItem(
            group=agg.group,
            date=agg.date,
            date_raw=agg.date_raw,
            pdf_date_normalised=agg.date.isoformat() if agg.date else "",
            source_amounts=[r.get(config.source.amount_field, "") for r in agg.source_records]
            if config.source.amount_field else [],
            aggregated_amount=agg.amount,
            invoice_breakup=agg.invoice_breakup,
            excel_formula_breakup=agg.excel_formula_breakup,
            return_amount=agg.return_amount,
            source_records=agg.source_records,
        )
        item.messages.extend(agg.warnings)

        # A row must have BOTH a valid (normalised) date and a valid amount to
        # enter the write plan.  This filters parser noise such as a title or
        # "From Date:" line that slipped through with text like "wise"/"sales".
        has_valid_date = agg.date is not None
        has_valid_amount = bool(agg.source_amount_values)
        if not has_valid_date or not has_valid_amount:
            item.status = Status.INVALID_ROW
            why = []
            if not has_valid_date:
                why.append(f"no valid date (got {agg.date_raw!r})")
            if not has_valid_amount:
                why.append("no valid amount")
            item.messages.append("Row skipped: " + "; ".join(why))
            report.skipped_invalid += 1
            report.items.append(item)
            continue

        sheet = config.group_to_sheet.get(agg.group, "")
        if not sheet:
            item.status = Status.UNMAPPED_GROUP
            item.messages.append("No worksheet mapped for this group.")
            report.items.append(item)
            if agg.group not in report.unmapped_groups:
                report.unmapped_groups.append(agg.group)
            continue
        item.sheet = sheet

        if sheet not in sheet_names:
            item.status = Status.SHEET_NOT_FOUND
            item.messages.append(f"Worksheet '{sheet}' not found in workbook.")
            report.errors += 1
            report.items.append(item)
            continue

        sm = config.sheets.get(sheet)
        if sm is None:
            item.status = Status.MISSING_COLUMN
            item.messages.append(f"No Excel column mapping configured for '{sheet}'.")
            report.missing_columns += 1
            report.items.append(item)
            continue

        # Resolve amount target.
        if sm.amount_target_mode == TargetMode.CELL_REFERENCE and sm.amount_cell:
            rc = excel_reader.cell_reference_to_rc(sm.amount_cell)
            if rc is None:
                item.status = Status.ERROR
                item.messages.append(f"Invalid target cell: {sm.amount_cell}")
                report.errors += 1
                report.items.append(item)
                continue
            item.matched_row, item.target_column = rc
            item.target_cell = sm.amount_cell.upper()
            item.target_column_letter = "".join(ch for ch in item.target_cell if ch.isalpha())
        else:
            col = excel_reader.resolve_column_index(
                wb, sheet,
                mode=sm.amount_target_mode.value,
                selector=sm.amount_column,
                header_row=sm.header_row,
            )
            if col is None:
                item.status = Status.MISSING_COLUMN
                item.messages.append(
                    f"Target amount column '{sm.amount_column}' not found in '{sheet}'."
                )
                report.missing_columns += 1
                report.items.append(item)
                continue
            item.target_column = col
            from openpyxl.utils import get_column_letter
            item.target_column_letter = get_column_letter(col)

            # Resolve the matching row by date.
            date_col = excel_reader.resolve_column_index(
                wb, sheet,
                mode=sm.date_target_mode.value,
                selector=sm.date_column,
                header_row=sm.header_row,
            )
            if date_col is None and sm.date_column:
                item.status = Status.MISSING_COLUMN
                item.messages.append(f"Date column '{sm.date_column}' not found in '{sheet}'.")
                report.missing_columns += 1
                report.items.append(item)
                continue

            matched = excel_reader.find_row_by_date(
                wb, sheet, date_col, agg.date,
                header_row=sm.header_row,
                date_formats=sm.date_formats or None,
                interpretation=interp,
            ) if date_col else None

            # Track match statistics per sheet (for the >25%-missing warning).
            per_sheet_targets.setdefault(sheet, set()).add(agg.date)
            if matched is None:
                per_sheet_missing[sheet] = per_sheet_missing.get(sheet, 0) + 1

            item.date_column = date_col
            if matched is None:
                # Diagnose WHY it's missing using a full-column scan.
                dset, mn, mx = _date_col_info(sheet, date_col, sm) if date_col else (set(), None, None)
                in_range = (mn is not None and mn <= agg.date <= mx)

                # If the exact date IS present in the column but find_row_by_date
                # did not match it, that's a genuine bug — surface it as an error
                # regardless of the missing-date behaviour.
                if agg.date in dset:
                    item.status = Status.ERROR
                    item.messages.append(
                        f"BUG: PDF date {item.pdf_date_normalised} exists in the Excel date "
                        f"column but was not matched. Please report this."
                    )
                    report.errors += 1
                    report.items.append(item)
                    continue

                action = config.write_rules.date_not_found_action
                do_insert = (action in (DateNotFoundAction.INSERT_ROW, DateNotFoundAction.COPY_NEAREST)
                             or config.write_rules.insert_missing_date_rows)
                if do_insert:
                    ws = wb[sheet]
                    base = _insert_position(wb, sheet, date_col, agg.date, sm)
                    cursor = append_cursor.get(sheet)
                    if base > ws.max_row:
                        # Append below existing data: distinct sequential rows.
                        row = cursor if cursor is not None else base
                        item.is_append = True
                        item.messages.append("Date not found; will append a new row.")
                    else:
                        # Mid-sheet insert: offset by inserts already planned above.
                        row = base if cursor is None else max(base, cursor)
                        item.messages.append("Date not found; will insert a new row.")
                    append_cursor[sheet] = row + 1
                    item.is_insert = True
                    item.matched_row = row
                    item.template_row = excel_reader.nearest_date_row(
                        wb, sheet, date_col, agg.date,
                        header_row=sm.header_row, date_formats=sm.date_formats or None,
                        interpretation=interp,
                    )
                elif in_range:
                    # Within the Excel date range but no exact row -> per the
                    # spec, flag as an error so nothing is written until resolved.
                    item.status = Status.ERROR
                    item.messages.append(
                        f"PDF date {item.pdf_date_normalised} is WITHIN the Excel date range "
                        f"[{mn} .. {mx}] but no row has this exact date. Likely a missing row "
                        f"or a date-format mismatch — resolve before writing."
                    )
                    report.errors += 1
                    report.items.append(item)
                    continue
                else:
                    item.status = Status.DATE_NOT_FOUND
                    rng = f"[{mn} .. {mx}]" if mn else "(no parseable dates)"
                    item.messages.append(
                        f"PDF date {item.pdf_date_normalised} is OUTSIDE the Excel date "
                        f"range {rng} for this column — likely the wrong worksheet or "
                        f"workbook year."
                    )
                    report.dates_not_found += 1
                    report.items.append(item)
                    continue
            else:
                item.matched_row = matched
                # Record the Excel date cell value + its normalised form.
                if date_col:
                    raw_excel = wb[sheet].cell(row=matched, column=date_col).value
                    item.excel_date_value = raw_excel
                    norm = utils.parse_date(raw_excel, formats=sm.date_formats or None,
                                            interpretation=interp)
                    item.excel_date_normalised = norm.isoformat() if norm else ""
            item.target_cell = f"{item.target_column_letter}{item.matched_row}"

        # Resolve optional return column.
        if config.source.return_field and (sm.return_column or sm.return_cell):
            if sm.return_target_mode == TargetMode.CELL_REFERENCE and sm.return_cell:
                rc = excel_reader.cell_reference_to_rc(sm.return_cell)
                if rc:
                    item.return_column = rc[1]
            else:
                item.return_column = excel_reader.resolve_column_index(
                    wb, sheet,
                    mode=sm.return_target_mode.value,
                    selector=sm.return_column,
                    header_row=sm.header_row,
                )

        # Compute existing value & final value.
        if not item.is_insert and item.matched_row and item.target_column:
            ws = wb[sheet]
            item.existing_value = ws.cell(row=item.matched_row, column=item.target_column).value
            if item.return_column:
                item.return_existing_value = ws.cell(row=item.matched_row, column=item.return_column).value

        item.final_value = _compute_final_value(item, agg, config)
        item.write_action = config.write_rules.write_action.value

        # Conflict detection.
        if (
            not item.is_insert
            and item.existing_value not in (None, "")
            and config.write_rules.write_action in (WriteAction.ASK, WriteAction.SKIP_IF_VALUE)
        ):
            item.status = Status.CONFLICT
            item.messages.append("Target cell already has a value.")
            report.conflicts += 1
        else:
            item.status = Status.READY

        report.items.append(item)

    # Strong warning when most PDF dates are missing from a sheet's date column.
    for sheet, targets in per_sheet_targets.items():
        total = len(targets)
        missing = per_sheet_missing.get(sheet, 0)
        if total and (missing / total) > DATE_MISS_WARN_THRESHOLD:
            sm = config.sheets.get(sheet)
            suggestions = []
            if sm is not None:
                try:
                    suggestions = excel_reader.suggest_date_columns(
                        wb, sheet, targets, header_row=sm.header_row,
                        date_formats=sm.date_formats or None, interpretation=interp)
                except Exception:
                    suggestions = []
            sug_txt = ""
            better = [s for s in suggestions if s[2] > (total - missing)]
            if better:
                sug_txt = " Suggested date column(s): " + ", ".join(
                    f"{letter} ({hdr or 'no header'}) matches {m}" for letter, hdr, m, _ in better)
            # Include the Excel column's actual date range vs the PDF range so a
            # "wrong year" workbook is obvious at a glance.
            rng_txt = ""
            date_col = excel_reader.resolve_column_index(
                wb, sheet, mode=sm.date_target_mode.value,
                selector=sm.date_column, header_row=sm.header_row) if sm else None
            if date_col:
                _, mn, mx = _date_col_info(sheet, date_col, sm)
                pdf_min, pdf_max = (min(targets), max(targets)) if targets else (None, None)
                if mn:
                    rng_txt = (f" Excel date range here is [{mn} .. {mx}]; "
                               f"PDF dates are [{pdf_min} .. {pdf_max}].")
            report.date_match_warnings.append(
                f"Sheet '{sheet}': {missing} of {total} PDF dates were not found in the "
                f"selected date column. This may mean the wrong date column, wrong worksheet, "
                f"wrong workbook year, or missing date rows.{rng_txt}{sug_txt}"
            )

    return report


def excel_date_debug_rows(wb, config: AppConfig) -> List[dict]:
    """Row-by-row debug of the resolved date column for every mapped sheet.

    Powers ``excel_date_debug.csv`` so the real stored value, data type and
    number format of each date cell are visible (independent of display).
    """
    interp = config.source.date_interpretation or "dmy"
    sheet_names = set(excel_reader.list_sheets(wb))
    out: List[dict] = []
    seen_sheets = set()
    for sheet in config.group_to_sheet.values():
        if not sheet or sheet in seen_sheets or sheet not in sheet_names:
            continue
        seen_sheets.add(sheet)
        sm = config.sheets.get(sheet)
        if sm is None:
            continue
        date_col = excel_reader.resolve_column_index(
            wb, sheet, mode=sm.date_target_mode.value,
            selector=sm.date_column, header_row=sm.header_row)
        if not date_col:
            continue
        out.extend(excel_reader.date_column_debug(
            wb, sheet, date_col, header_row=sm.header_row,
            date_formats=sm.date_formats or None, interpretation=interp))
    return out


NO_RANGE_MESSAGE = "The selected worksheet does not contain the PDF date range."


def date_diagnostics(wb, config: AppConfig, rows) -> List[dict]:
    """Per-mapped-worksheet date diagnostic for the Page 7 panel.

    For each worksheet that has at least one Customer Name mapped to it, returns:
    worksheet, customers, selected date column, Excel min/max date (full scan),
    first/last 10 parsed dates, the PDF date range for the customers mapped here,
    matched/missing counts, and a plain-language status.
    """
    interp = config.source.date_interpretation or "dmy"
    sheet_names = set(excel_reader.list_sheets(wb))

    # Collect the PDF dates per target worksheet (only valid, normalised dates).
    per_sheet: Dict[str, dict] = {}
    for agg in rows:
        if agg.date is None:
            continue
        sheet = config.group_to_sheet.get(agg.group, "")
        if not sheet:
            continue
        d = per_sheet.setdefault(sheet, {"dates": set(), "groups": set()})
        d["dates"].add(agg.date)
        d["groups"].add(agg.group)

    out: List[dict] = []
    for sheet, info in per_sheet.items():
        pdf_dates = info["dates"]
        row = {
            "worksheet": sheet,
            "customers": sorted(info["groups"]),
            "date_column": "",
            "excel_min": None,
            "excel_max": None,
            "first10": [],
            "last10": [],
            "pdf_min": min(pdf_dates) if pdf_dates else None,
            "pdf_max": max(pdf_dates) if pdf_dates else None,
            "matched": 0,
            "missing": len(pdf_dates),
            "parsed_count": 0,
            "status": "",
        }
        sm = config.sheets.get(sheet)
        if sheet not in sheet_names:
            row["status"] = f"Worksheet '{sheet}' not found in the uploaded workbook."
            out.append(row); continue
        if sm is None:
            row["status"] = "No column mapping configured for this worksheet."
            out.append(row); continue
        date_col = excel_reader.resolve_column_index(
            wb, sheet, mode=sm.date_target_mode.value,
            selector=sm.date_column, header_row=sm.header_row)
        row["date_column"] = sm.date_column
        if not date_col:
            row["status"] = f"Date column '{sm.date_column}' not found in '{sheet}'."
            out.append(row); continue

        an = excel_reader.analyze_date_column(
            wb, sheet, date_col, pdf_dates, header_row=sm.header_row,
            date_formats=sm.date_formats or None, interpretation=interp)
        row.update({
            "excel_min": an["min_date"], "excel_max": an["max_date"],
            "first10": an["first10"], "last10": an["last10"],
            "matched": an["matched"], "missing": an["missing"],
            "parsed_count": an["parsed_count"],
        })

        covers = (an["min_date"] is not None and row["pdf_min"] is not None
                  and an["min_date"] <= row["pdf_min"] and row["pdf_max"] <= an["max_date"])
        if an["parsed_count"] == 0:
            row["status"] = "No parseable dates found in the selected date column."
        elif not covers:
            row["status"] = NO_RANGE_MESSAGE
        elif an["missing"] > 0:
            # Worksheet covers the PDF range but some dates still unmatched -> bug.
            row["status"] = ("BUG: worksheet covers the PDF date range but "
                             f"{an['missing']} date(s) were not matched.")
        else:
            row["status"] = "OK"
        out.append(row)
    return out


def _insert_position(wb, sheet, date_col, target_date, sm) -> int:
    """Choose a row index to insert at: keep dates sorted when possible."""
    ws = wb[sheet]
    if not date_col or target_date is None:
        return ws.max_row + 1
    insert_at = ws.max_row + 1
    for r in range(sm.header_row + 1, ws.max_row + 1):
        val = ws.cell(row=r, column=date_col).value
        parsed = utils.parse_date(val, formats=sm.date_formats or None)
        if parsed is not None and parsed > target_date:
            insert_at = r
            break
    return insert_at


def _compute_final_value(item: PlanItem, agg: AggregatedRow, config: AppConfig):
    """Pick the value to write based on the chosen output mode.

    The numeric total (``item.aggregated_amount``) is always preserved on the
    plan item; this only decides what lands in the target cell.
    """
    if config.write_rules.output_type == OutputType.FORMULA:
        # Use the precomputed comma-free formula, e.g. "=630+12705".
        return agg.excel_formula_breakup or build_formula([agg.amount])
    if config.write_rules.write_action == WriteAction.ADD:
        base = item.existing_value if isinstance(item.existing_value, (int, float)) else 0
        return round((base or 0) + agg.amount, 4)
    return round(agg.amount, 4)


def plan_to_write_ops(report: ValidationReport, config: AppConfig) -> List[WriteOp]:
    """Convert READY (and conflict-resolved) plan items into write operations."""
    ops: List[WriteOp] = []
    for item in report.items:
        if item.status not in (Status.READY, Status.CONFLICT):
            continue
        if item.target_column is None or item.matched_row is None:
            continue
        comment = ""
        if config.write_rules.add_source_comment and item.source_records:
            lines = [str(r.get("_source_text", "")) for r in item.source_records]
            comment = "Source rows:\n" + "\n".join(l for l in lines if l)

        # For inserted/appended rows, also stamp the date into the date column so
        # the new row is identifiable. Whichever op for this row comes first
        # performs the physical insert / formatting copy.
        date_op_added = False
        if item.is_insert and item.date_column and item.date is not None:
            ops.append(WriteOp(
                sheet_name=item.sheet,
                row=item.matched_row,
                column=item.date_column,
                value=item.date,
                action=WriteAction.REPLACE,
                output_type=OutputType.NUMERIC,
                insert_row=item.is_insert and not item.is_append,
                append_only=item.is_append,
                template_row=item.template_row,
                label=f"{item.group} / {item.date_raw} (date)",
            ))
            date_op_added = True

        ops.append(WriteOp(
            sheet_name=item.sheet,
            row=item.matched_row,
            column=item.target_column,
            value=item.final_value,
            action=config.write_rules.write_action,
            output_type=config.write_rules.output_type,
            comment=comment,
            # If a date op already inserted/formatted this row, the amount op just
            # writes its cell; otherwise the amount op performs the insert/append.
            insert_row=item.is_insert and not item.is_append and not date_op_added,
            append_only=item.is_append and not date_op_added,
            template_row=item.template_row,
            label=f"{item.group} / {item.date_raw}",
        ))

        # Optional return write.
        if item.return_column and config.source.return_field:
            ops.append(WriteOp(
                sheet_name=item.sheet,
                row=item.matched_row,
                column=item.return_column,
                value=round(item.return_amount, 4),
                action=config.write_rules.write_action,
                output_type=OutputType.NUMERIC,
                label=f"{item.group} / {item.date_raw} (return)",
            ))
    return ops


def plan_to_export_rows(report: ValidationReport, config: AppConfig) -> List[dict]:
    """Flatten the write plan into export rows (for write_plan.csv).

    Always includes the numeric ``aggregated_amount``, the human-readable
    ``invoice_breakup`` and the ``excel_formula_breakup`` so all three value
    forms are visible regardless of the chosen write mode.
    """
    mode = config.write_rules.output_type.value
    out: List[dict] = []
    for it in report.items:
        out.append({
            "customer_name": it.group,
            "worksheet": it.sheet,
            "date": it.date_raw,
            "num_source_rows": len(it.source_records),
            "invoice_breakup": it.invoice_breakup,
            "aggregated_amount": round(it.aggregated_amount, 2),
            "excel_formula_breakup": it.excel_formula_breakup,
            "return_amount": round(it.return_amount, 2),
            "matched_row": it.matched_row,
            "target_column": it.target_column_letter,
            "target_cell": it.target_cell,
            "existing_value": it.existing_value,
            "write_mode": mode,
            "value_to_write": it.final_value,
            "status": it.status,
            "messages": "; ".join(it.messages),
        })
    return out
