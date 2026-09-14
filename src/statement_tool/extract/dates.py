"""Best-effort normalisation of the many date formats banks use."""
from __future__ import annotations

import re
from datetime import datetime

_MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

_FORMATS = [
    "%d/%m/%Y", "%d/%m/%y",
    "%d-%m-%Y", "%d-%m-%y",
    "%d %b %Y", "%d %B %Y",
    "%Y-%m-%d", "%Y/%m/%d",
    "%m/%d/%Y",
]


def parse_date(raw: str | None) -> str | None:
    """Return an ISO 'YYYY-MM-DD' string, or None if unparsable.

    Kept deliberately permissive (many formats tried) since a wrong guess
    here is caught later - a downstream unparsable/blank date makes the row
    show up in the run summary rather than silently sorting wrong.
    """
    if not raw:
        return None
    text = re.sub(r"\s+", " ", str(raw)).strip()
    if not text:
        return None

    for fmt in _FORMATS:
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue

    m = re.match(r"^(\d{1,2})\s+([A-Za-z]{3,})\.?\s+(\d{2,4})$", text)
    if m:
        day, mon_text, year = m.groups()
        mon = _MONTHS.get(mon_text.lower()[:3])
        if mon:
            year_num = int(year)
            if year_num < 100:
                year_num += 2000
            try:
                return datetime(year_num, mon, int(day)).date().isoformat()
            except ValueError:
                return None

    return None
