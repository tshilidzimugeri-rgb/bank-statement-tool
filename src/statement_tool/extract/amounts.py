"""Parsing of bank-statement-flavoured numbers into floats."""
from __future__ import annotations

import re


def parse_amount(raw: str | None) -> float | None:
    """Parse strings like 'R 1,234.56', '(1,234.56)', '1 234,56 Dr', '-'.

    Returns None for blank/placeholder cells (e.g. '-', '', 'n/a') rather
    than 0.0, so a genuinely empty debit/credit cell doesn't get summed as a
    zero-value transaction.
    """
    if raw is None:
        return None
    text = str(raw).strip()
    if not text or text in {"-", "--", "n/a", "na"}:
        return None

    is_negative = False
    if text.startswith("(") and text.rstrip().endswith(")"):
        is_negative = True

    lower = text.lower()
    if lower.rstrip().endswith("dr"):
        is_negative = True

    # Strip currency symbols, parentheses, Cr/Dr suffixes, and whitespace.
    cleaned = re.sub(r"(?i)\b(cr|dr)\b", "", text)
    cleaned = cleaned.replace("R", "").replace("r", "")
    cleaned = cleaned.strip(" ()")
    cleaned = cleaned.replace(" ", "")

    if not cleaned:
        return None

    # Handle both "1,234.56" (comma thousands) and "1.234,56" (dot thousands,
    # comma decimal) styles, since a config might encounter either.
    if "," in cleaned and "." in cleaned:
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        # Ambiguous: "1,234" (thousands) vs "1234,56" (decimal comma).
        # Treat a trailing 2-digit group as decimals, else as thousands sep.
        if re.match(r"^-?\d+,\d{2}$", cleaned):
            cleaned = cleaned.replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")

    try:
        value = float(cleaned)
    except ValueError:
        return None

    if is_negative:
        value = -abs(value)
    return value
