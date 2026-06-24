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

# A broad, configurable set of date formats.  The user can extend this through
# the mapping configuration; we never assume one particular format.
DEFAULT_DATE_FORMATS: Sequence[str] = (
    "%d/%m/%Y",
    "%d-%m-%Y",
    "%d.%m.%Y",
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%m/%d/%Y",
    "%m-%d-%Y",
    "%d/%m/%y",
    "%d-%m-%y",
    "%d %b %Y",
    "%d %B %Y",
    "%b %d, %Y",
    "%B %d, %Y",
    "%d-%b-%Y",
    "%d-%b-%y",
)


def parse_date(
    value,
    *,
    formats: Optional[Iterable[str]] = None,
    dayfirst: bool = True,
) -> Optional[_dt.date]:
    """Parse a date from many possible representations.

    Accepts real ``date``/``datetime`` objects (as produced by openpyxl for
    real Excel dates) as well as text in any of the supplied ``formats``.
    Returns ``None`` when nothing matches.
    """
    if value is None:
        return None
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value

    text = str(value).strip()
    if not text:
        return None

    fmts = list(formats) if formats else list(DEFAULT_DATE_FORMATS)
    for fmt in fmts:
        try:
            return _dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    # Last resort: pull a date-looking token out of a longer string and retry
    # with a numeric heuristic that respects ``dayfirst``.
    match = re.search(r"(\d{1,4})[/\-.](\d{1,2})[/\-.](\d{1,4})", text)
    if match:
        a, b, c = (int(g) for g in match.groups())
        return _heuristic_numeric_date(a, b, c, dayfirst=dayfirst)
    return None


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
