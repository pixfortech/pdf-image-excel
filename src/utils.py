"""Generic helper utilities for parsing amounts and dates.

Nothing in this module is specific to any particular business, customer,
worksheet, column, or document layout.  Everything is driven by the values
that are passed in (or by sensible, configurable defaults).
"""
from __future__ import annotations

import datetime as _dt
import re
from typing import Iterable, Optional, Sequence

# ---------------------------------------------------------------------------
# Amount parsing
# ---------------------------------------------------------------------------

# Characters that commonly decorate a numeric amount.  These are stripped
# before the number is parsed.  We deliberately keep this configurable rather
# than hardcoding a single currency.
_DEFAULT_CURRENCY_CHARS = "₹$€£¥"


def parse_amount(
    value,
    *,
    decimal_sep: str = ".",
    thousands_sep: str = ",",
    currency_chars: str = _DEFAULT_CURRENCY_CHARS,
) -> Optional[float]:
    """Parse a human written amount into a float.

    Handles commas, currency symbols, surrounding spaces, parentheses for
    negatives and an optional trailing/leading minus sign.  Returns ``None``
    when the value cannot be interpreted as a number.

    The separators are configurable so locales that use ``.`` as a thousands
    separator and ``,`` as a decimal separator are supported too.
    """
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    negative = False
    # Parentheses denote negative numbers in many accounting reports.
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1].strip()

    # Strip currency symbols and spaces.
    for ch in currency_chars:
        text = text.replace(ch, "")
    text = text.replace(" ", "").replace(" ", "")

    if text.endswith("-"):
        negative = True
        text = text[:-1]
    if text.startswith("-"):
        negative = True
        text = text[1:]
    if text.startswith("+"):
        text = text[1:]

    # Normalise separators to a canonical ``1234.56`` form.
    if thousands_sep:
        text = text.replace(thousands_sep, "")
    if decimal_sep and decimal_sep != ".":
        text = text.replace(decimal_sep, ".")

    if text in ("", ".", "-"):
        return None

    if not re.fullmatch(r"[0-9]*\.?[0-9]*", text):
        # Keep only the leading numeric portion if extra characters slipped in.
        match = re.search(r"[0-9]*\.?[0-9]+", text)
        if not match:
            return None
        text = match.group(0)

    try:
        number = float(text)
    except ValueError:
        return None
    return -number if negative else number


def format_amount(value: Optional[float], decimals: int = 2) -> str:
    """Format a numeric amount for display with grouping separators."""
    if value is None:
        return ""
    return f"{value:,.{decimals}f}"


# ---------------------------------------------------------------------------
# Date parsing
# ---------------------------------------------------------------------------

# Date-interpretation modes used to disambiguate numeric dates like 05/06/2026.
DMY = "dmy"   # Indian / British: day first (DEFAULT)
MDY = "mdy"   # US: month first
AUTO = "auto"  # auto-detect from the components
DATE_INTERPRETATIONS = (DMY, MDY, AUTO)

# Unambiguous formats tried first regardless of interpretation.
_ISO_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d")
_NAMED_FORMATS = (
    "%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%d-%B-%Y", "%d %b %y", "%d-%b-%y",
    "%b %d, %Y", "%B %d, %Y", "%b %d %Y", "%B %d %Y",
)
# Day-first (Indian/British) numeric formats.
_DMY_FORMATS = (
    "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y", "%d.%m.%y",
)
# Month-first (US) numeric formats.
_MDY_FORMATS = (
    "%m/%d/%Y", "%m-%d-%Y", "%m.%d.%Y", "%m/%d/%y", "%m-%d-%y", "%m.%d.%y",
)

# Kept for backwards compatibility; a broad day-first-leaning set.
DEFAULT_DATE_FORMATS: Sequence[str] = _ISO_FORMATS + _NAMED_FORMATS + _DMY_FORMATS + _MDY_FORMATS

# Excel serial-date epoch (openpyxl/Excel "1900 system"; day 0 = 1899-12-30).
_EXCEL_EPOCH = _dt.date(1899, 12, 30)
# Only treat bare numbers in this range as Excel serial dates (≈ 1954-08-05 to
# ≈ 3543), so a stray year like 2026 or a small count is not misread as a date.
_SERIAL_MIN, _SERIAL_MAX = 20000, 600000


def _formats_for(interpretation: str) -> tuple:
    """Ordered format list for an interpretation (unambiguous first)."""
    interp = (interpretation or DMY).lower()
    if interp == MDY:
        return _ISO_FORMATS + _NAMED_FORMATS + _MDY_FORMATS + _DMY_FORMATS
    # DMY (default) and AUTO both lean day-first for the strptime pass; AUTO then
    # also runs the numeric heuristic below.
    return _ISO_FORMATS + _NAMED_FORMATS + _DMY_FORMATS + _MDY_FORMATS


def excel_serial_to_date(value) -> Optional[_dt.date]:
    """Convert an Excel serial date number to a ``date`` (or ``None``)."""
    if isinstance(value, bool):
        return None
    if not isinstance(value, (int, float)):
        return None
    if not (_SERIAL_MIN <= value <= _SERIAL_MAX):
        return None
    try:
        return _EXCEL_EPOCH + _dt.timedelta(days=int(round(value)))
    except (OverflowError, ValueError):
        return None


def parse_date(
    value,
    *,
    formats: Optional[Iterable[str]] = None,
    interpretation: str = DMY,
    allow_serial: bool = True,
    dayfirst: Optional[bool] = None,
) -> Optional[_dt.date]:
    """Parse a date from many representations and normalise to a ``date``.

    Handles real ``date``/``datetime`` objects (as produced by openpyxl for real
    Excel dates), Excel serial-date numbers, and text in many formats (DD/MM/YYYY,
    DD-MM-YYYY, DD.MM.YYYY, YYYY-MM-DD, ``15 May 2026``, ``15-May-2026``,
    two-digit years, datetime strings like ``2026-05-15 00:00:00``).

    ``interpretation`` (``dmy``/``mdy``/``auto``) disambiguates numeric dates such
    as ``05/06/2026``.  The default is ``dmy`` (Indian/British → 5 June 2026).
    ``formats`` (if given) are tried first as an explicit user hint.  Always
    returns a date with no time component, or ``None`` when nothing matches.
    """
    if value is None:
        return None
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value

    # Excel serial date numbers (when a date cell isn't formatted as a date).
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if allow_serial:
            return excel_serial_to_date(value)
        return None

    # ``dayfirst`` retained for backwards compatibility.
    if dayfirst is False and interpretation == DMY:
        interpretation = MDY

    text = str(value).strip()
    if not text:
        return None
    # Drop a trailing time component, e.g. "2026-05-15 00:00:00" -> "2026-05-15".
    text_nodate_time = re.sub(r"[ T]\d{1,2}:\d{2}(:\d{2})?(\.\d+)?$", "", text).strip()

    candidates = []
    if formats:
        candidates.extend(formats)
    candidates.extend(_formats_for(interpretation))

    for source in (text_nodate_time, text):
        for fmt in candidates:
            try:
                return _dt.datetime.strptime(source, fmt).date()
            except ValueError:
                continue

    # Numeric heuristic for anything left (and the primary path for AUTO).
    match = re.search(r"(\d{1,4})[/\-.](\d{1,2})[/\-.](\d{1,4})", text)
    if match:
        a, b, c = (int(g) for g in match.groups())
        df = (interpretation != MDY)
        return _heuristic_numeric_date(a, b, c, dayfirst=df)
    return None


def normalise_date(value, **kwargs) -> Optional[_dt.date]:
    """Alias for :func:`parse_date` emphasising normalisation to a bare date."""
    return parse_date(value, **kwargs)


def _heuristic_numeric_date(a: int, b: int, c: int, *, dayfirst: bool) -> Optional[_dt.date]:
    """Best-effort interpretation of three numeric date components."""
    # Detect a 4-digit year position first.
    if a > 31:  # YYYY M D
        year, month, day = a, b, c
    elif c > 31:  # D M YYYY or M D YYYY
        year = c
        if dayfirst:
            day, month = a, b
        else:
            month, day = a, b
    else:  # Two digit year at the end.
        year = 2000 + c if c < 100 else c
        if dayfirst:
            day, month = a, b
        else:
            month, day = a, b

    # Swap if month/day are obviously transposed.
    if month > 12 and day <= 12:
        day, month = month, day
    try:
        return _dt.date(year, month, day)
    except ValueError:
        return None


def date_to_str(value: Optional[_dt.date], fmt: str = "%d/%m/%Y") -> str:
    if value is None:
        return ""
    return value.strftime(fmt)


# ---------------------------------------------------------------------------
# Misc helpers
# ---------------------------------------------------------------------------

def normalise_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def is_blank(value) -> bool:
    return value is None or str(value).strip() == ""
