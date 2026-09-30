"""The whole workflow as two calls, shared by the app and the scripts.

``prepare``: extract -> parse -> recognise profile -> validate fields ->
reconcile -> aggregate -> plan.  Nothing is written.

``update``:  apply the Ready cells to the uploaded workbook, verify that only
those cells changed, and remember the profile as the last successful one.
"""
from __future__ import annotations

import hashlib
from dataclasses import dataclass, field
from typing import List, Optional

from . import extractor, parser, plan as plan_mod, profiles, workbook
from .profiles import Profile, ProfileStore

LOADED = "loaded"        # a saved profile recognised this report
NEW = "new"              # no saved profile: a suggested one needs confirming
GIVEN = "given"          # the caller supplied the profile explicitly


@dataclass
class Prepared:
    pdf_name: str
    xlsx_name: str
    original: bytes                      # untouched uploaded workbook (backup)
    document: extractor.Document
    report: parser.ParsedReport
    profile: Profile
    profile_source: str
    field_issues: dict = field(default_factory=dict)
    invoices: List[parser.Invoice] = field(default_factory=list)
    rejected: List[parser.IgnoredLine] = field(default_factory=list)
    reconciliation: Optional[parser.Reconciliation] = None
    totals: List[plan_mod.DailyTotal] = field(default_factory=list)
    plan: plan_mod.Plan = field(default_factory=plan_mod.Plan)
    wb: Optional[workbook.Workbook] = None
    workbook_error: str = ""

    # ---- readiness -------------------------------------------------------
    @property
    def hard_blockers(self) -> List[str]:
        """Problems that no acknowledgement can override."""
        out = []
        if not self.report.groups:
            out.append("No groups were found in the PDF.")
        if self.field_issues:
            out.extend(self.field_issues.values())
        elif not self.invoices:
            out.append("No valid invoice rows (with a valid date and amount) were extracted.")
        if self.reconciliation and not self.reconciliation.ok:
            for line in self.reconciliation.failures:
                out.append(f"Totals do not reconcile for {line.scope}: extracted "
                           f"{line.extracted:,.2f} vs printed {line.printed:,.2f} "
                           f"(difference {line.difference:,.2f}).")
        if self.workbook_error:
            out.append(self.workbook_error)
        return out

    @property
    def ready_count(self) -> int:
        return len(self.plan.ready)

    def can_update(self, exclude_blocked: bool = False) -> bool:
        if self.hard_blockers or self.ready_count == 0:
            return False
        return exclude_blocked or not self.plan.blocking


@dataclass
class UpdateResult:
    workbook: bytes
    changes: List[workbook.CellChange]
    rows: List[plan_mod.PlanRow]


class UpdateError(Exception):
    pass


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()[:16]


def prepare(pdf_bytes: bytes, pdf_name: str, xlsx_bytes: bytes, xlsx_name: str,
            store: Optional[ProfileStore] = None, profile: Optional[Profile] = None) -> Prepared:
    doc = extractor.extract(pdf_bytes, pdf_name)

    if profile is not None:
        source = GIVEN
    else:
        probe = parser.parse(doc)
        profile = store.find_for(probe) if store else None
        source = LOADED if profile else NEW
        if profile is None:
            profile = profiles.suggest_profile(probe)

    s = profile.source
    report = parser.parse(doc, group_label=s.group_label, header_labels=s.columns,
                          group_total_label=s.group_total_label, date_order=s.date_order)
    prepared = Prepared(pdf_name, xlsx_name, xlsx_bytes, doc, report, profile, source)
    prepared.field_issues = profiles.field_issues(profile, report)

    if not prepared.field_issues:
        prepared.invoices, prepared.rejected = parser.invoices(
            report, date_field=s.date_field, amount_field=s.amount_field,
            returns_field=s.returns_field, reference_field=s.reference_field)
        prepared.reconciliation = parser.reconcile(report, prepared.invoices)
        prepared.totals = plan_mod.aggregate(prepared.invoices)

    try:
        prepared.wb = workbook.Workbook(xlsx_bytes)
    except workbook.WorkbookError as exc:
        prepared.workbook_error = str(exc)
        return prepared
    prepared.plan = plan_mod.build_plan(prepared.totals, profile, prepared.wb)
    return prepared


def update(prepared: Prepared, store: Optional[ProfileStore] = None,
           exclude_blocked: bool = False) -> UpdateResult:
    """Write the Ready cells into the uploaded workbook and verify the result."""
    if not prepared.can_update(exclude_blocked):
        if prepared.ready_count == 0 and not prepared.hard_blockers:
            raise UpdateError("Nothing is ready to write.")
        raise UpdateError("The workbook cannot be updated until the listed problems are resolved.")
    changes = prepared.plan.changes()
    updated = workbook.update_workbook(prepared.original, changes)
    problems = workbook.verify_update(prepared.original, updated, changes)
    if problems:
        raise UpdateError("Verification failed; the workbook was not updated: " + "; ".join(problems))

    if store is not None:
        # Remember names that matched through a space-insensitive alias.
        for status in prepared.plan.sheets:
            if status.note.startswith("matched saved name"):
                profiles.remember_group(prepared.profile, status.group, status.sheet)
        store.mark_success(prepared.profile)
    return UpdateResult(updated, changes, prepared.plan.ready)
