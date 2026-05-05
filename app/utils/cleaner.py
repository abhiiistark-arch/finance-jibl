"""String and numeric normalization for extracted cells."""

from __future__ import annotations

import re
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Optional, Union


def normalize_header(text: object) -> str:
    """Lowercase and remove spaces, slashes, and non-alphanumeric for fuzzy compare."""
    if text is None:
        return ""
    s = str(text).strip().lower()
    s = re.sub(r"[^\w\s]", "", s)
    s = re.sub(r"\s+", "", s)
    return s.strip()


_WHITESPACE_RE = re.compile(r"\s+", re.UNICODE)
_COMMA_RE = re.compile(r"[,\u00a0\u202f]'?")


def clean_cell_string(value: object) -> str:
    """Strip and collapse whitespace; empty string if None."""
    if value is None:
        return ""
    s = _WHITESPACE_RE.sub(" ", str(value)).strip()
    return s


def parse_numeric(value: object) -> Optional[Union[float, int]]:
    """
    Strip commas/thousand separators and parse number safely.
    Returns None if not parseable.
    """
    if value is None:
        return None
    s = clean_cell_string(value)
    if not s or s.lower() in ("-", "—", "n/a", "na"):
        return None
    s = _COMMA_RE.sub("", s).strip()
    s = re.sub(r"^\(?(.+)\)?$", r"-\1", s) if s.startswith("(") and s.endswith(")") else s
    try:
        d = Decimal(s)
        if d == int(d):
            return int(d)
        return float(d)
    except (InvalidOperation, ValueError):
        pass
    try:
        return float(s)
    except ValueError:
        return None


def format_excel_numeric(value: object) -> str:
    """Store numbers in Excel-compatible string form without thousands commas."""
    n = parse_numeric(value)
    if n is None:
        return ""
    return str(int(n)) if isinstance(n, int) else str(n)


_MONTH_ABBR = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


def format_risk_end_date(value: object) -> str:
    """
    Display policy expiry / risk end dates as dd-Mon-yyyy (e.g. 26-Mar-2027).
    Accepts dd/mm/yyyy variants from PDF extraction; ignores locale for month abbreviations.
    """
    raw = clean_cell_string(value)
    if not raw:
        return ""

    dept_heading = normalize_header("Department Name")
    if normalize_header(raw) == dept_heading:
        return ""

    s = raw.split()[0].strip()

    canon_m = re.match(r"^(\d{1,2})[\s\-/]([A-Za-z]{3})[\s\-/](\d{4})$", s)
    if canon_m:
        d, mon_txt, y = canon_m.groups()
        mon_lower = mon_txt.lower()
        for abbrev in _MONTH_ABBR:
            if abbrev.lower() == mon_lower:
                return f"{int(d)}-{abbrev}-{int(y)}"

    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y", "%d.%m.%y"):
        try:
            dt = datetime.strptime(s, fmt)
            mon = _MONTH_ABBR[dt.month - 1]
            return f"{dt.day}-{mon}-{dt.year}"
        except ValueError:
            continue

    return raw


def blank_if_nan(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and (value != value):  # NaN
        return ""
    return str(value) if value else ""
