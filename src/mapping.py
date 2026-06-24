"""Configuration data structures for source/PDF mapping, Excel mapping,
group-to-worksheet mapping and write rules.

Every business rule lives in data, not in code.  These dataclasses are
serialisable to/from JSON so a user can save a configuration once and reuse it
for the same PDF/Excel format later.  No customer name, worksheet name, column
name/letter, header text, or date format is hardcoded here.
"""
from __future__ import annotations

import dataclasses
import datetime as _dt
import json
from dataclasses import dataclass, field
from enum import Enum
from typing import Dict, List, Optional


# ---------------------------------------------------------------------------
# Enumerations describing user choices
# ---------------------------------------------------------------------------

class TargetMode(str, Enum):
    """How an Excel target column/cell is identified."""
    HEADER_NAME = "header_name"
    COLUMN_LETTER = "column_letter"
    CELL_REFERENCE = "cell_reference"


class WriteAction(str, Enum):
    """What to do when writing to an existing cell."""
    REPLACE = "replace"
    ADD = "add"
    SKIP_IF_VALUE = "skip_if_value"
    ASK = "ask"


class OutputType(str, Enum):
    """Whether to write a numeric total or an Excel formula."""
    NUMERIC = "numeric"
    FORMULA = "formula"


class DateNotFoundAction(str, Enum):
    SKIP = "skip"
    INSERT_ROW = "insert_row"
    COPY_NEAREST = "copy_nearest"
    MANUAL = "manual"


# ---------------------------------------------------------------------------
# Source / PDF mapping
# ---------------------------------------------------------------------------

@dataclass
class SourceMapping:
    """Maps detected source fields to their semantic role.

    Field names here refer to *detected* field keys (e.g. column headers found
    in the PDF, or generic ``col_0`` style keys).  They are chosen by the user
    in the UI, never assumed by the code.
    """
    group_field: str = ""           # field that represents customer/branch/party/...
    date_field: str = ""            # field that represents the row date
    amount_field: str = ""          # field that represents the main amount
    return_field: str = ""          # optional deduction/return field
    ignored_fields: List[str] = field(default_factory=list)

    sum_duplicate_dates: bool = True
    keep_invoice_breakup: bool = True
    output_type: OutputType = OutputType.NUMERIC

    # Parsing rules (configurable, not hardcoded).
    date_formats: List[str] = field(default_factory=list)
    decimal_sep: str = "."
    thousands_sep: str = ","
    currency_chars: str = "₹$€£¥"

    # The label text used in the PDF to introduce a group section, e.g.
    # "Customer Name". Configurable; the parser also auto-detects candidates.
    group_label: str = ""


# ---------------------------------------------------------------------------
# Excel mapping (per worksheet)
# ---------------------------------------------------------------------------

@dataclass
class SheetMapping:
    """Describes how to write into a single worksheet."""
    sheet_name: str = ""
    header_row: int = 1

    # Date column used to locate the matching row.
    date_target_mode: TargetMode = TargetMode.HEADER_NAME
    date_column: str = ""           # header name or column letter

    # Amount target.
    amount_target_mode: TargetMode = TargetMode.HEADER_NAME
    amount_column: str = ""         # header / letter / cell ref depending on mode

    # Optional return/deduction target.
    return_target_mode: TargetMode = TargetMode.HEADER_NAME
    return_column: str = ""

    # Optional fixed target cell (overrides row matching when set).
    amount_cell: str = ""
    return_cell: str = ""

    # Date format used when the worksheet stores dates as text.
    date_formats: List[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Write rules
# ---------------------------------------------------------------------------

@dataclass
class WriteRules:
    write_action: WriteAction = WriteAction.ASK
    output_type: OutputType = OutputType.NUMERIC
    insert_missing_date_rows: bool = False
    date_not_found_action: DateNotFoundAction = DateNotFoundAction.SKIP
    add_source_comment: bool = False
    aggregation_keys: List[str] = field(default_factory=lambda: ["group", "date"])


# ---------------------------------------------------------------------------
# Top level configuration bundle
# ---------------------------------------------------------------------------

@dataclass
class AppConfig:
    source: SourceMapping = field(default_factory=SourceMapping)
    # Customer/group -> worksheet name. Filled in entirely by the user.
    group_to_sheet: Dict[str, str] = field(default_factory=dict)
    # worksheet name -> SheetMapping
    sheets: Dict[str, SheetMapping] = field(default_factory=dict)
    write_rules: WriteRules = field(default_factory=WriteRules)

    created: str = ""
    note: str = ""


# ---------------------------------------------------------------------------
# Mapping-pattern helpers
# ---------------------------------------------------------------------------

def copy_sheet_mapping(template: SheetMapping, sheet_name: str) -> SheetMapping:
    """Return a copy of ``template``'s column/date pattern for ``sheet_name``.

    Everything except the worksheet name is copied verbatim, so the same column
    layout (date column, amount/return targets, header row, target modes, date
    formats) is reused.  Cell-reference targets are intentionally cleared on the
    copy because an exact cell reference (e.g. ``F10``) only makes sense for the
    one sheet it was authored on.
    """
    copy = dataclasses.replace(template, sheet_name=sheet_name)
    copy.date_formats = list(template.date_formats)
    # Exact cell references are sheet-specific; drop them on copies so the copied
    # pattern resolves by column instead.
    copy.amount_cell = ""
    copy.return_cell = ""
    if copy.amount_target_mode == TargetMode.CELL_REFERENCE:
        copy.amount_target_mode = TargetMode.COLUMN_LETTER
    if copy.return_target_mode == TargetMode.CELL_REFERENCE:
        copy.return_target_mode = TargetMode.COLUMN_LETTER
    return copy


def apply_pattern_to_sheets(
    config: "AppConfig",
    template_sheet: str,
    target_sheets,
    *,
    overwrite: bool = True,
) -> List[str]:
    """Copy the template sheet's column pattern onto every target sheet.

    ``target_sheets`` is typically every worksheet assigned to a Customer Name
    group (``config.group_to_sheet.values()``).  Returns the list of sheet names
    that were written.  When ``overwrite`` is False, sheets that already have a
    mapping are left untouched (useful to preserve manual overrides).
    """
    template = config.sheets.get(template_sheet)
    if template is None:
        return []
    written: List[str] = []
    for sheet in dict.fromkeys(target_sheets):  # de-dup, preserve order
        if not sheet:
            continue
        if sheet == template_sheet:
            continue
        if not overwrite and sheet in config.sheets:
            continue
        config.sheets[sheet] = copy_sheet_mapping(template, sheet)
        written.append(sheet)
    return written


# ---------------------------------------------------------------------------
# (De)serialisation helpers
# ---------------------------------------------------------------------------

def _to_jsonable(obj):
    if dataclasses.is_dataclass(obj):
        return {k: _to_jsonable(v) for k, v in dataclasses.asdict(obj).items()}
    if isinstance(obj, Enum):
        return obj.value
    if isinstance(obj, dict):
        return {k: _to_jsonable(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_to_jsonable(v) for v in obj]
    if isinstance(obj, (_dt.date, _dt.datetime)):
        return obj.isoformat()
    return obj


def config_to_dict(config: AppConfig) -> dict:
    return _to_jsonable(config)


def config_to_json(config: AppConfig, indent: int = 2) -> str:
    return json.dumps(config_to_dict(config), indent=indent, ensure_ascii=False)


def save_config(config: AppConfig, path: str) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(config_to_json(config))


def _enum_from(value, enum_cls, default):
    if value is None:
        return default
    try:
        return enum_cls(value)
    except ValueError:
        return default


def source_from_dict(data: dict) -> SourceMapping:
    data = data or {}
    sm = SourceMapping()
    for f in dataclasses.fields(sm):
        if f.name in data:
            setattr(sm, f.name, data[f.name])
    sm.output_type = _enum_from(data.get("output_type"), OutputType, OutputType.NUMERIC)
    return sm


def sheet_from_dict(data: dict) -> SheetMapping:
    data = data or {}
    sm = SheetMapping()
    for f in dataclasses.fields(sm):
        if f.name in data:
            setattr(sm, f.name, data[f.name])
    sm.date_target_mode = _enum_from(data.get("date_target_mode"), TargetMode, TargetMode.HEADER_NAME)
    sm.amount_target_mode = _enum_from(data.get("amount_target_mode"), TargetMode, TargetMode.HEADER_NAME)
    sm.return_target_mode = _enum_from(data.get("return_target_mode"), TargetMode, TargetMode.HEADER_NAME)
    return sm


def write_rules_from_dict(data: dict) -> WriteRules:
    data = data or {}
    wr = WriteRules()
    for f in dataclasses.fields(wr):
        if f.name in data:
            setattr(wr, f.name, data[f.name])
    wr.write_action = _enum_from(data.get("write_action"), WriteAction, WriteAction.ASK)
    wr.output_type = _enum_from(data.get("output_type"), OutputType, OutputType.NUMERIC)
    wr.date_not_found_action = _enum_from(
        data.get("date_not_found_action"), DateNotFoundAction, DateNotFoundAction.SKIP
    )
    return wr


def config_from_dict(data: dict) -> AppConfig:
    data = data or {}
    cfg = AppConfig()
    cfg.source = source_from_dict(data.get("source", {}))
    cfg.group_to_sheet = dict(data.get("group_to_sheet", {}) or {})
    cfg.sheets = {
        name: sheet_from_dict(sd) for name, sd in (data.get("sheets", {}) or {}).items()
    }
    cfg.write_rules = write_rules_from_dict(data.get("write_rules", {}))
    cfg.created = data.get("created", "")
    cfg.note = data.get("note", "")
    return cfg


def config_from_json(text: str) -> AppConfig:
    return config_from_dict(json.loads(text))


def load_config(path: str) -> AppConfig:
    with open(path, "r", encoding="utf-8") as fh:
        return config_from_json(fh.read())
