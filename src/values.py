"""Parsing of amounts, dates and names.

Every rule for turning text into a number, a date or a lookup key lives here,
so the parser, the profile and the workbook code all agree.  Parsing is STRICT:
a value must *be* an amount or a date, not merely contain one — this is what
stops an invoice number such as ``ABC-I12345`` from ever becoming a date.
"""
from __future__ import annotations

import datetime as _dt
import re
from typing import Iterable, Optional

DMY = "dmy"    # Indian / British: 05/06/2026 = 5 June 2026 (default)
MDY = "mdy"    # US: 05/06/2026 = 6 May 2026
AUTO = "auto"  # decide per document from unambiguous dates
DATE_ORDERS = (DMY, MDY, AUTO)

_CURRENCY = "₹$€£¥"
_AMOUNT_RE = re.compile(r"^[+-]?(?:\d{1,3}(?:,\d{2,3})*|\d+)(?:\.\d+)?$")
_NUMERIC_DATE_RE = re.compile(r"^(\d{1,4})([/.\-])(\d{1,2})\2(\d{1,4})$")
_TIME_SUFFIX_RE = re.compile(r"[ T]\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?$")
_NAMED_FORMATS = (
    "%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%d-%B-%Y", "%d %b %y", "%d-%b-%y",
    "%b %d, %Y", "%B %d, %Y", "%b %d %Y", "%B %d %Y",
)
_EXCEL_EPOCH = _dt.date(1899, 12, 30)
_SERIAL_RANGE = (20000, 600000)  # ~1954..3543; keeps small numbers/years out


# ---------------------------------------------------------------------------
# Amounts
# ---------------------------------------------------------------------------

def parse_amount(value) -> Optional[float]:
    """Parse ``8,405.00``, ``₹ 1,23,456.78``, ``(500.00)`` or a number.

    Returns ``None`` for anything that is not entirely an amount.
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    for ch in _CURRENCY:
        text = text.replace(ch, "")
    text = text.replace(" ", "").replace(" ", "")
    if not text:
        return None
    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative, text = True, text[1:-1]
    if text.endswith("-"):
        negative, text = True, text[:-1]
    if not _AMOUNT_RE.match(text):
        return None
    number = float(text.replace(",", ""))
    return -abs(number) if negative else number


def is_amount(value) -> bool:
    return parse_amount(value) is not None


# ---------------------------------------------------------------------------
# Dates
# ---------------------------------------------------------------------------

def parse_date(value, order: str = DMY, *, allow_serial: bool = False) -> Optional[_dt.date]:
    """Strictly parse a date and normalise it to a bare ``date``.

    Accepts ``date``/``datetime`` objects (openpyxl reads real Excel dates this
    way regardless of display format such as ``d-mmm``), Excel serial numbers
    when ``allow_serial``, and text that is *entirely* a date: ``28/08/2026``,
    ``28-08-2026``, ``28.08.26``, ``2026-08-28``, ``2026-08-28 00:00:00``,
    ``28 Aug 2026``, ``28-Aug-2026``.  ``order`` resolves ambiguous numeric
    dates (``dmy`` default).
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    if isinstance(value, (int, float)):
        if allow_serial and _SERIAL_RANGE[0] <= value <= _SERIAL_RANGE[1]:
            return _EXCEL_EPOCH + _dt.timedelta(days=int(round(value)))
        return None

    text = _TIME_SUFFIX_RE.sub("", str(value).strip()).strip()
    if not text:
        return None

    m = _NUMERIC_DATE_RE.match(text)
    if m:
        a, _sep, b, c = m.groups()
        if len(a) == 4:                       # YYYY-MM-DD
            return _make(int(a), int(b), int(c))
        if len(c) not in (2, 4):
            return None
        year = int(c) + (2000 if len(c) == 2 else 0)
        first, second = int(a), int(b)
        if order == MDY:
            return _make(year, first, second)
        return _make(year, second, first)     # DMY and AUTO (resolved upstream)

    for fmt in _NAMED_FORMATS:
        try:
            return _dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _make(year: int, month: int, day: int) -> Optional[_dt.date]:
    try:
        return _dt.date(year, month, day)
    except ValueError:
        return None


def is_date(value, order: str = DMY) -> bool:
    return parse_date(value, order) is not None


def detect_date_order(texts: Iterable[str], default: str = DMY) -> str:
    """Resolve ``auto`` for a document: a first component > 12 proves
    day-first, a second component > 12 proves month-first.  Conflicting or no
    evidence falls back to ``default``."""
    dmy = mdy = False
    for t in texts:
        m = _NUMERIC_DATE_RE.match(str(t).strip())
        if not m or len(m.group(1)) == 4:
            continue
        first, second = int(m.group(1)), int(m.group(3))
        dmy |= first > 12
        mdy |= second > 12
    if dmy and not mdy:
        return DMY
    if mdy and not dmy:
        return MDY
    return default


def format_date(d: Optional[_dt.date]) -> str:
    return d.strftime("%d/%m/%Y") if d else ""


# ---------------------------------------------------------------------------
# Keys used for matching (never for display)
# ---------------------------------------------------------------------------

def label_key(text) -> str:
    """Key for column/field labels: ``Inv No`` == ``InvNo`` == ``inv. no``."""
    return re.sub(r"[^a-z0-9]", "", str(text or "").lower())


def name_key(text) -> str:
    """Key for group names: case, repeated/edge spaces and punctuation ignored,
    but word boundaries kept (``NORTH MARKET`` != ``NORTHMARKET`` at this level)."""
    cleaned = re.sub(r"[^\w\s&]", " ", str(text or "").upper())
    return re.sub(r"\s+", " ", cleaned).strip()


def compact_key(text) -> str:
    """Space-insensitive group key: ``NORTH MARKET`` == ``NORTHMARKET``."""
    return name_key(text).replace(" ", "")


def format_number(value) -> str:
    """Display helper: 19750.0 -> '19,750.00'."""
    if value is None or value == "":
        return ""
    if isinstance(value, (int, float)):
        return f"{value:,.2f}"
    return str(value)
