"""Grouping and aggregation of parsed records into a write plan.

The aggregation key is fully configurable (group+date, group+date+field,
date-only, or no aggregation).  Amounts and dates are parsed using the rules in
the :class:`~src.mapping.SourceMapping`.  No customer, date, amount, or column
is hardcoded.
"""
from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence

from . import utils
from .mapping import SourceMapping
from .parser import Record


@dataclass
class AggregatedRow:
    group: str
    date: Optional[_dt.date]
    date_raw: str
    amount: float
    return_amount: float
    source_row_indexes: List[int] = field(default_factory=list)
    source_records: List[dict] = field(default_factory=list)  # invoice breakup
    warnings: List[str] = field(default_factory=list)
    extra_key: str = ""   # value of the optional extra aggregation field


def _get_field(rec: Record, field_name: str) -> str:
    if not field_name:
        return ""
    return rec.fields.get(field_name, "")


def aggregate(
    records: Sequence[Record],
    source: SourceMapping,
    *,
    aggregation_keys: Optional[Sequence[str]] = None,
) -> List[AggregatedRow]:
    """Aggregate ``records`` according to the source mapping and keys.

    ``aggregation_keys`` is a list drawn from {"group", "date", "<field>"}.
    When it contains no date or the source mapping says not to sum duplicate
    dates, rows are still produced but each source row stays separate unless the
    keys collapse them.
    """
    keys = list(aggregation_keys) if aggregation_keys is not None else ["group", "date"]
    formats = source.date_formats or None

    buckets: Dict[tuple, AggregatedRow] = {}
    order: List[tuple] = []

    for rec in records:
        if rec.ignored:
            continue
        group = rec.group or _get_field(rec, source.group_field)
        date_raw = _get_field(rec, source.date_field)
        date_val = utils.parse_date(date_raw, formats=formats)
        amount_raw = _get_field(rec, source.amount_field)
        amount = utils.parse_amount(
            amount_raw,
            decimal_sep=source.decimal_sep,
            thousands_sep=source.thousands_sep,
            currency_chars=source.currency_chars,
        )
        return_raw = _get_field(rec, source.return_field)
        return_amount = utils.parse_amount(
            return_raw,
            decimal_sep=source.decimal_sep,
            thousands_sep=source.thousands_sep,
            currency_chars=source.currency_chars,
        ) if source.return_field else None

        warnings: List[str] = []
        if source.amount_field and amount is None:
            warnings.append(f"Could not parse amount: {amount_raw!r}")
            amount = 0.0
        if source.date_field and date_raw and date_val is None:
            warnings.append(f"Could not parse date: {date_raw!r}")

        extra_key = ""
        key_parts: List = []
        for k in keys:
            if k == "group":
                key_parts.append(("group", group))
            elif k == "date":
                key_parts.append(("date", date_val.isoformat() if date_val else date_raw))
            else:
                val = _get_field(rec, k)
                extra_key = val
                key_parts.append((k, val))
        # When no aggregation keys are provided, keep each row separate.
        if not keys:
            key_parts.append(("row", rec.row_index))
        key = tuple(key_parts)

        if key not in buckets or not source.sum_duplicate_dates:
            if key in buckets and not source.sum_duplicate_dates:
                # Make a unique key so rows are not merged.
                key = key + (("_row", rec.row_index),)
            agg = AggregatedRow(
                group=group,
                date=date_val,
                date_raw=date_raw,
                amount=amount or 0.0,
                return_amount=return_amount or 0.0,
                source_row_indexes=[rec.row_index],
                source_records=[dict(rec.fields, _page=rec.page, _source_text=rec.source_text)],
                warnings=list(warnings),
                extra_key=extra_key,
            )
            buckets[key] = agg
            order.append(key)
        else:
            agg = buckets[key]
            agg.amount += amount or 0.0
            if return_amount:
                agg.return_amount += return_amount
            agg.source_row_indexes.append(rec.row_index)
            agg.source_records.append(dict(rec.fields, _page=rec.page, _source_text=rec.source_text))
            agg.warnings.extend(warnings)

    return [buckets[k] for k in order]


def group_totals(rows: Sequence[AggregatedRow]) -> Dict[str, float]:
    totals: Dict[str, float] = {}
    for r in rows:
        totals[r.group] = totals.get(r.group, 0.0) + r.amount
    return totals


def build_formula(amounts: Sequence[float]) -> str:
    """Build an Excel SUM-style formula showing the source breakup."""
    parts = "+".join(f"{a:g}" for a in amounts)
    return f"={parts}" if parts else "=0"
