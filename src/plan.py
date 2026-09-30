"""Aggregate invoices and plan exactly which cells will change.

One status vocabulary for the whole app.  A plan row is one target cell for
one (group, date, field).  Only ``Ready`` rows are written.
"""
from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from . import values
from .parser import Invoice
from .profiles import FORMULA, SKIP_NONEMPTY, Profile, lookup_group
from .layout import resolve_layout
from .workbook import CellChange, Workbook

READY_NEW = "Ready"
READY_OVERWRITE = "Ready (overwrites)"
UNCHANGED = "Unchanged"
KEEP_EXISTING = "Skipped (cell has a value)"
IGNORED = "Ignored group"
UNMAPPED = "Not mapped"
AMBIGUOUS = "Ambiguous mapping"
SHEET_MISSING = "Worksheet missing"
COLUMN_PROBLEM = "Column problem"
DATE_NOT_FOUND = "Date not found"
DUPLICATE_DATE = "Date on several rows"
FORMULA_CELL = "Target is a formula"
MERGED_CELL = "Target is merged"

READY = {READY_NEW, READY_OVERWRITE}
BLOCKING = {UNMAPPED, AMBIGUOUS, SHEET_MISSING, COLUMN_PROBLEM, DATE_NOT_FOUND,
            DUPLICATE_DATE, FORMULA_CELL, MERGED_CELL}


@dataclass
class DailyTotal:
    """All invoices of one group on one date (e.g. 31,535 + 15,518 = 19,750)."""
    group: str
    date: _dt.date
    invoices: List[Invoice] = field(default_factory=list)

    @property
    def amount(self) -> float:
        return round(sum(i.amount or 0 for i in self.invoices), 2)

    @property
    def amounts(self) -> List[float]:
        return [i.amount for i in self.invoices if i.amount is not None]

    @property
    def returns(self) -> Optional[float]:
        """``None`` when no invoice carries a Returns value (blank != zero)."""
        vals = [i.returns for i in self.invoices if i.returns is not None]
        return round(sum(vals), 2) if vals else None

    @property
    def breakup(self) -> str:
        return " + ".join(values.format_number(a) for a in self.amounts)

    @property
    def formula(self) -> str:
        parts = [_formula_number(a) for a in self.amounts] or ["0"]
        text = parts[0] + "".join(p if p.startswith("-") else "+" + p for p in parts[1:])
        return "=" + text

    @property
    def references(self) -> str:
        return ", ".join(i.reference for i in self.invoices if i.reference)


def _formula_number(v: float) -> str:
    return str(int(v)) if float(v).is_integer() else repr(float(v))


def aggregate(invoices: List[Invoice]) -> List[DailyTotal]:
    """Group by (group, date) and sum — order: first appearance of group, then date."""
    buckets: Dict[tuple, DailyTotal] = {}
    for inv in invoices:
        buckets.setdefault((inv.group, inv.date), DailyTotal(inv.group, inv.date)).invoices.append(inv)
    order = list(dict.fromkeys(i.group for i in invoices))
    return sorted(buckets.values(), key=lambda t: (order.index(t.group), t.date))


@dataclass
class PlanRow:
    total: DailyTotal
    field: str                 # "amount" | "returns"
    sheet: str = ""
    cell: str = ""
    existing: object = None
    new_value: object = None
    status: str = ""
    note: str = ""

    @property
    def is_ready(self) -> bool:
        return self.status in READY

    @property
    def is_blocking(self) -> bool:
        return self.status in BLOCKING


@dataclass
class SheetStatus:
    group: str
    sheet: str
    dates_matched: int
    dates_total: int
    target: str
    status: str
    note: str = ""
    via: str = ""              # "pattern" (common pattern) or "exception" (sheet override)
    returns_target: str = ""


@dataclass
class Plan:
    rows: List[PlanRow] = field(default_factory=list)
    sheets: List[SheetStatus] = field(default_factory=list)

    @property
    def ready(self) -> List[PlanRow]:
        return [r for r in self.rows if r.is_ready]

    @property
    def blocking(self) -> List[PlanRow]:
        return [r for r in self.rows if r.is_blocking]

    def changes(self) -> List[CellChange]:
        return [CellChange(r.sheet, r.cell, r.new_value) for r in self.ready]


def build_plan(totals: List[DailyTotal], profile: Profile, wb: Workbook) -> Plan:
    plan = Plan()
    settings = profile.write
    by_group: Dict[str, List[DailyTotal]] = {}
    for t in totals:
        by_group.setdefault(t.group, []).append(t)

    for group, group_totals in by_group.items():
        match = lookup_group(profile, group)
        sheet = match.sheet
        group_status, note, resolved = "", "", None
        date_letter = amount_letter = returns_letter = None
        # A returns target is needed only when there is a return value to write
        # (blank source returns are never written unless zero-fill is on).
        need_returns = bool(settings.write_returns and profile.source.returns_field
                            and (settings.zero_fill_returns
                                 or any(t.returns is not None for t in group_totals)))

        if match.status == "ignored":
            group_status = IGNORED
        elif match.status == "ambiguous":
            group_status, note = AMBIGUOUS, "could be " + " or ".join(match.candidates)
        elif match.status == "unmapped":
            group_status = UNMAPPED
        elif not wb.has_sheet(sheet):
            group_status, note = SHEET_MISSING, f"'{sheet}' is not in this workbook"
        elif not profile.layout_for(sheet).amount.is_set:
            group_status, note = COLUMN_PROBLEM, "worksheet pattern not set up yet"
        else:
            resolved = resolve_layout(wb, sheet, profile.layout_for(sheet), need_returns)
            if resolved.ok:
                date_letter, amount_letter, returns_letter = resolved.date, resolved.amount, resolved.returns
            else:
                group_status, note = COLUMN_PROBLEM, "; ".join(resolved.problems)

        dates_index = wb.date_rows(sheet, date_letter, resolved.header_row, profile.source.date_order) \
            if date_letter else {}
        matched = sum(1 for t in group_totals if len(dates_index.get(t.date, [])) == 1)
        if match.status == "alias" and not note:
            note = f"matched saved name '{match.saved_name}'"
        plan.sheets.append(SheetStatus(
            group, sheet, matched, len(group_totals),
            resolved.describe(amount_letter) if amount_letter else "",
            group_status or _group_ready(matched, group_totals), note,
            via="exception" if sheet in profile.sheet_overrides else ("pattern" if sheet else ""),
            returns_target=resolved.describe(returns_letter) if returns_letter else ""))

        for t in group_totals:
            targets = [("amount", amount_letter)]
            if returns_letter and (t.returns is not None or settings.zero_fill_returns):
                targets.append(("returns", returns_letter))
            for field_name, letter in targets:
                row = PlanRow(t, field_name, sheet=sheet)
                plan.rows.append(row)
                if group_status:
                    row.status, row.note = group_status, note
                    continue
                rows = dates_index.get(t.date, [])
                if not rows:
                    row.status, row.note = DATE_NOT_FOUND, f"{values.format_date(t.date)} not in {sheet}"
                    continue
                if len(rows) > 1:
                    row.status, row.note = DUPLICATE_DATE, f"rows {', '.join(map(str, rows))}"
                    continue
                row.cell = f"{letter}{rows[0]}"
                row.new_value = _new_value(t, field_name, settings.output_mode)
                _decide(row, wb.cell(sheet, row.cell), settings.existing_values)
    return plan


def _group_ready(matched: int, totals: List[DailyTotal]) -> str:
    return READY_NEW if matched == len(totals) else DATE_NOT_FOUND


def _new_value(t: DailyTotal, field_name: str, output_mode: str):
    if field_name == "returns":
        return t.returns if t.returns is not None else 0.0
    if output_mode == FORMULA and len(t.amounts) > 1:
        return t.formula
    return t.amount


def _decide(row: PlanRow, state, existing_policy: str) -> None:
    row.existing = state.value
    if state.kind == "formula" and not is_breakup_formula(state.value):
        row.status, row.note = FORMULA_CELL, f"contains {state.value}"
    elif state.kind == "merged":
        row.status = MERGED_CELL
    elif state.kind == "empty":
        row.status = READY_NEW
    elif _same(state.value, row.new_value):
        row.status = UNCHANGED
    elif existing_policy == SKIP_NONEMPTY:
        row.status = KEEP_EXISTING
    else:
        row.status = READY_OVERWRITE


def is_breakup_formula(value) -> bool:
    """``=12500+7250`` — numbers only, as written by formula output mode."""
    return isinstance(value, str) and re.fullmatch(r"=[\d.]+(?:[+-][\d.]+)*", value.strip()) is not None


def _same(existing, new) -> bool:
    if isinstance(new, str) or isinstance(existing, str):
        return str(existing).strip() == str(new)
    return isinstance(existing, (int, float)) and not isinstance(existing, bool) \
        and abs(float(existing) - float(new)) < 0.005
