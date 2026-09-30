"""Reusable mapping profiles: SOURCE SCHEMA + TARGET BEHAVIOUR.

A profile records, for one kind of report (e.g. "Customer Wise Sales"):

* the source schema — group label, printed column labels, and which column is
  the date / amount / returns / reference field (by LABEL, never by position);
* the target behaviour — Customer Name -> worksheet, a worksheet column
  template plus per-sheet overrides, and write settings.

Profiles are user configuration.  They live in a local store outside the
repository (``~/.pdf-image-excel/profiles`` or ``%APPDATA%``) and are never
committed.  Nothing in this module knows any business name.
"""
from __future__ import annotations

import datetime as _dt
import json
import os
import re
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

from . import values

PROFILE_VERSION = 2

NUMERIC = "numeric"
FORMULA = "formula"
OVERWRITE = "overwrite"
SKIP_NONEMPTY = "skip_nonempty"


# ---------------------------------------------------------------------------
# Schema
# ---------------------------------------------------------------------------

@dataclass
class ColumnRef:
    """A worksheet column: by letter (validated against the saved header text)
    or, when no letter is stored, by a header that must be unique."""
    letter: str = ""
    header: str = ""

    @property
    def is_set(self) -> bool:
        return bool(self.letter or self.header)

    def describe(self) -> str:
        if self.letter and self.header:
            return f"{self.letter} ({self.header})"
        return self.letter or self.header


@dataclass
class SheetLayout:
    header_row: int = 1
    date: ColumnRef = field(default_factory=ColumnRef)
    amount: ColumnRef = field(default_factory=ColumnRef)
    returns: ColumnRef = field(default_factory=ColumnRef)


@dataclass
class SourceSchema:
    group_label: str = ""
    columns: List[str] = field(default_factory=list)
    date_field: str = ""
    amount_field: str = ""
    returns_field: str = ""
    reference_field: str = ""
    group_total_label: str = ""
    date_order: str = values.DMY


@dataclass
class WriteSettings:
    output_mode: str = NUMERIC            # numeric total, or "=a+b" formula breakup
    existing_values: str = OVERWRITE      # or skip_nonempty
    write_returns: bool = True            # only when a real Returns value exists
    zero_fill_returns: bool = False       # blank source returns -> write 0 (opt-in)


@dataclass
class Profile:
    profile_name: str = "New profile"
    source: SourceSchema = field(default_factory=SourceSchema)
    group_to_sheet: Dict[str, str] = field(default_factory=dict)
    ignored_groups: List[str] = field(default_factory=list)
    sheet_template: SheetLayout = field(default_factory=SheetLayout)
    sheet_overrides: Dict[str, SheetLayout] = field(default_factory=dict)
    write: WriteSettings = field(default_factory=WriteSettings)
    profile_version: int = PROFILE_VERSION
    created: str = ""
    updated: str = ""
    last_success: str = ""

    def layout_for(self, sheet: str) -> SheetLayout:
        return self.sheet_overrides.get(sheet, self.sheet_template)

    def mapped_sheets(self) -> List[str]:
        return list(dict.fromkeys(s for s in self.group_to_sheet.values() if s))


# ---------------------------------------------------------------------------
# Group lookup with safe alias normalisation
# ---------------------------------------------------------------------------

@dataclass
class GroupMatch:
    status: str                    # mapped | alias | ignored | ambiguous | unmapped
    sheet: str = ""
    saved_name: str = ""
    candidates: List[str] = field(default_factory=list)


def lookup_group(profile: Profile, name: str) -> GroupMatch:
    """Find the worksheet for a group name.

    Matches exactly after normalising case, punctuation and repeated spaces;
    then space-insensitively (``NORTH MARKET`` == ``NORTHMARKET``) only if that is
    unambiguous.  No fuzzy matching.
    """
    key, compact = values.name_key(name), values.compact_key(name)
    if any(values.name_key(g) == key for g in profile.ignored_groups):
        return GroupMatch("ignored")
    exact = [g for g in profile.group_to_sheet if values.name_key(g) == key]
    sheets = {profile.group_to_sheet[g] for g in exact}
    if len(sheets) == 1:
        return GroupMatch("mapped", sheets.pop(), exact[0])
    if len(sheets) > 1:
        return GroupMatch("ambiguous", candidates=sorted(sheets))
    if any(values.compact_key(g) == compact for g in profile.ignored_groups):
        return GroupMatch("ignored")
    near = [g for g in profile.group_to_sheet if values.compact_key(g) == compact]
    sheets = {profile.group_to_sheet[g] for g in near}
    if len(sheets) == 1:
        return GroupMatch("alias", sheets.pop(), near[0])
    if len(sheets) > 1:
        return GroupMatch("ambiguous", candidates=sorted(sheets))
    return GroupMatch("unmapped")


def remember_group(profile: Profile, name: str, sheet: str) -> None:
    """Record ``name -> sheet`` (also used to settle an ambiguity once)."""
    key = values.name_key(name)
    for g in [g for g in profile.group_to_sheet if values.name_key(g) == key]:
        del profile.group_to_sheet[g]
    profile.ignored_groups = [g for g in profile.ignored_groups if values.name_key(g) != key]
    profile.group_to_sheet[name] = sheet


def ignore_group(profile: Profile, name: str) -> None:
    key = values.name_key(name)
    for g in [g for g in profile.group_to_sheet if values.name_key(g) == key]:
        del profile.group_to_sheet[g]
    if not any(values.name_key(g) == key for g in profile.ignored_groups):
        profile.ignored_groups.append(name)


# ---------------------------------------------------------------------------
# Source compatibility: restore fields by meaning, never by position
# ---------------------------------------------------------------------------

_EXPECTED_KIND = {
    "date_field": ("date",),
    "amount_field": ("amount",),
    "returns_field": ("amount", "empty"),
    "reference_field": ("text", "amount", "date", "empty"),
}
FIELD_NAMES = {
    "date_field": "Date", "amount_field": "Main amount",
    "returns_field": "Returns", "reference_field": "Reference",
}


def field_issues(profile: Profile, report) -> Dict[str, str]:
    """For each saved field: is the SAME labelled column present in this
    document, and does its data have the expected type?  A failing field is
    reported for re-mapping; another column is never substituted."""
    issues: Dict[str, str] = {}
    for attr, kinds in _EXPECTED_KIND.items():
        label = getattr(profile.source, attr)
        required = attr in ("date_field", "amount_field")
        if not label:
            if required:
                issues[attr] = f"{FIELD_NAMES[attr]} field is not mapped."
            continue
        if not report.has_column(label):
            issues[attr] = (f"Saved {FIELD_NAMES[attr].lower()} column '{label}' is not in this "
                            f"document (columns found: {', '.join(report.columns) or 'none'}).")
            continue
        kind = report.column_kind(label)
        if kind not in kinds:
            issues[attr] = (f"Column '{label}' was expected to contain "
                            f"{' or '.join(kinds)} values but contains {kind} values.")
    return issues


def source_matches(profile: Profile, report) -> bool:
    """Structural recognition: same group label and the date/amount columns
    exist.  Filenames are never used."""
    s = profile.source
    return bool(
        s.group_label and s.date_field and s.amount_field
        and values.label_key(s.group_label) == values.label_key(report.group_label)
        and report.has_column(s.date_field) and report.has_column(s.amount_field)
    )


def suggest_profile(report, name: str = "") -> Profile:
    """A starting profile for a report with no saved profile yet.  Fields are
    suggested from column data types; the user confirms them once."""
    kinds = {c: report.column_kind(c) for c in report.columns}
    dates = [c for c, k in kinds.items() if k == "date"]
    amounts = [c for c, k in kinds.items() if k == "amount"]
    texts = [c for c, k in kinds.items() if k == "text"]
    returns = [c for c, k in kinds.items()
               if k in ("amount", "empty") and "return" in values.label_key(c)]
    amount = next((c for c in amounts if c not in returns), "")
    return Profile(
        profile_name=name or (f"{report.group_label} report" if report.group_label else "New profile"),
        source=SourceSchema(
            group_label=report.group_label,
            columns=list(report.columns),
            date_field=dates[0] if len(dates) == 1 else "",
            amount_field=amount,
            returns_field=returns[0] if returns else "",
            reference_field=texts[0] if texts else "",
            group_total_label=report.group_total_label,
            date_order=report.date_order,
        ),
    )


# ---------------------------------------------------------------------------
# (De)serialisation
# ---------------------------------------------------------------------------

def to_dict(profile: Profile) -> dict:
    return asdict(profile)


def to_json(profile: Profile) -> str:
    return json.dumps(to_dict(profile), indent=2, ensure_ascii=False)


def _ref(d) -> ColumnRef:
    d = d or {}
    return ColumnRef(letter=str(d.get("letter", "")).upper(), header=str(d.get("header", "")))


def _layout(d) -> SheetLayout:
    d = d or {}
    return SheetLayout(int(d.get("header_row", 1) or 1), _ref(d.get("date")),
                       _ref(d.get("amount")), _ref(d.get("returns")))


def from_dict(data: dict) -> Profile:
    if not isinstance(data, dict):
        raise ValueError("A profile must be a JSON object.")
    if "profile_version" not in data:
        return _from_legacy(data)
    if int(data["profile_version"]) > PROFILE_VERSION:
        raise ValueError(f"Profile version {data['profile_version']} is newer than this app supports.")
    src = data.get("source", {}) or {}
    wr = data.get("write", {}) or {}
    return Profile(
        profile_name=data.get("profile_name", "Imported profile"),
        source=SourceSchema(**{k: src[k] for k in SourceSchema.__dataclass_fields__ if k in src}),
        group_to_sheet=dict(data.get("group_to_sheet", {}) or {}),
        ignored_groups=list(data.get("ignored_groups", []) or []),
        sheet_template=_layout(data.get("sheet_template")),
        sheet_overrides={k: _layout(v) for k, v in (data.get("sheet_overrides", {}) or {}).items()},
        write=WriteSettings(**{k: wr[k] for k in WriteSettings.__dataclass_fields__ if k in wr}),
        created=data.get("created", ""), updated=data.get("updated", ""),
        last_success=data.get("last_success", ""),
    )


def _from_legacy(data: dict) -> Profile:
    """Import a mapping JSON saved by the previous (v1) app.

    Keeps the valuable parts — group->sheet assignments, worksheet columns and
    write settings.  Positional source fields (``col_1``...) are dropped on
    purpose; the date/amount fields are then re-mapped by label.
    """
    src = data.get("source", {}) or {}

    def label(v):
        return "" if not v or re.fullmatch(r"col_\d+", str(v)) or v == "group" else str(v)

    def legacy_ref(mode, selector):
        selector = str(selector or "")
        if not selector:
            return ColumnRef()
        return ColumnRef(letter=selector.upper()) if mode == "column_letter" else ColumnRef(header=selector)

    layouts = {}
    for name, sm in (data.get("sheets", {}) or {}).items():
        layouts[name] = SheetLayout(
            int(sm.get("header_row", 1) or 1),
            legacy_ref(sm.get("date_target_mode"), sm.get("date_column")),
            legacy_ref(sm.get("amount_target_mode"), sm.get("amount_column")),
            legacy_ref(sm.get("return_target_mode"), sm.get("return_column")),
        )
    template = next(iter(layouts.values()), SheetLayout())
    rules = data.get("write_rules", {}) or {}
    return Profile(
        profile_name="Imported mapping",
        source=SourceSchema(group_label=src.get("group_label", ""),
                            date_field=label(src.get("date_field")),
                            amount_field=label(src.get("amount_field")),
                            returns_field=label(src.get("return_field"))),
        group_to_sheet=dict(data.get("group_to_sheet", {}) or {}),
        sheet_template=template,
        sheet_overrides={k: v for k, v in layouts.items() if v != template},
        write=WriteSettings(
            output_mode=FORMULA if rules.get("output_type") == "formula" else NUMERIC,
            existing_values=OVERWRITE if rules.get("write_action") == "replace" else SKIP_NONEMPTY,
        ),
    )


# ---------------------------------------------------------------------------
# Local store
# ---------------------------------------------------------------------------

def default_store_dir() -> Path:
    if os.environ.get("PDFX_PROFILE_DIR"):
        return Path(os.environ["PDFX_PROFILE_DIR"])
    if os.name == "nt" and os.environ.get("APPDATA"):
        return Path(os.environ["APPDATA"]) / "pdf-image-excel" / "profiles"
    return Path.home() / ".pdf-image-excel" / "profiles"


def _now() -> str:
    return _dt.datetime.now().isoformat(timespec="seconds")


def _slug(name: str) -> str:
    slug = re.sub(r"[^A-Za-z0-9]+", "-", name).strip("-").lower()
    return slug or "profile"


class ProfileStore:
    """One JSON file per profile in a local directory."""

    def __init__(self, root: Optional[Path] = None):
        self.root = Path(root) if root else default_store_dir()
        self.root.mkdir(parents=True, exist_ok=True)

    def _path(self, name: str) -> Path:
        return self.root / f"{_slug(name)}.json"

    def names(self) -> List[str]:
        return sorted(p.profile_name for p in self.all())

    def all(self) -> List[Profile]:
        out = []
        for path in sorted(self.root.glob("*.json")):
            try:
                out.append(from_dict(json.loads(path.read_text(encoding="utf-8"))))
            except (ValueError, json.JSONDecodeError, TypeError):
                continue
        return out

    def load(self, name: str) -> Optional[Profile]:
        path = self._path(name)
        if not path.exists():
            return None
        return from_dict(json.loads(path.read_text(encoding="utf-8")))

    def save(self, profile: Profile) -> Profile:
        profile.created = profile.created or _now()
        profile.updated = _now()
        self._path(profile.profile_name).write_text(to_json(profile), encoding="utf-8")
        return profile

    def delete(self, name: str) -> bool:
        path = self._path(name)
        if path.exists():
            path.unlink()
            return True
        return False

    def rename(self, old: str, new: str) -> Profile:
        profile = self.load(old)
        if profile is None:
            raise KeyError(old)
        if _slug(new) != _slug(old) and self._path(new).exists():
            raise ValueError(f"A profile named '{new}' already exists.")
        self.delete(old)
        profile.profile_name = new
        return self.save(profile)

    def import_json(self, text: str) -> Profile:
        profile = from_dict(json.loads(text))
        base, n = profile.profile_name, 2
        while self._path(profile.profile_name).exists():
            profile.profile_name = f"{base} ({n})"
            n += 1
        return self.save(profile)

    def mark_success(self, profile: Profile) -> Profile:
        profile.last_success = _now()
        return self.save(profile)

    def find_for(self, report) -> Optional[Profile]:
        """The most recently successful profile whose source structure matches."""
        matching = [p for p in self.all() if source_matches(p, report)]
        matching.sort(key=lambda p: (p.last_success, p.updated), reverse=True)
        return matching[0] if matching else None

    def last_successful(self) -> Optional[Profile]:
        used = [p for p in self.all() if p.last_success]
        return max(used, key=lambda p: p.last_success) if used else None
