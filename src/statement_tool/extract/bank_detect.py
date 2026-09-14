"""Figure out which bank a statement is from, and who the client is."""
from __future__ import annotations

import re

from ..config import BankLayout, ClientRule


def detect_bank(first_page_text: str, layouts: dict[str, BankLayout], generic: BankLayout) -> BankLayout:
    text_lower = (first_page_text or "").lower()
    for layout in layouts.values():
        if any(needle in text_lower for needle in layout.detect):
            return layout
    return generic


def extract_statement_period(full_text: str, layout: BankLayout, generic: BankLayout) -> str | None:
    text = full_text or ""
    for pattern in list(layout.period_patterns) + list(generic.period_patterns):
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            return re.sub(r"\s+", " ", m.group(1)).strip(" .:-")
    return None


def match_client(
    rules: list[ClientRule],
    *,
    sender: str | None = None,
    subject: str | None = None,
    filename: str | None = None,
    first_page_text: str | None = None,
) -> str | None:
    haystacks = [h.lower() for h in (sender, subject, filename, first_page_text) if h]
    for rule in rules:
        for needle in rule.match:
            if any(needle in haystack for haystack in haystacks):
                return rule.name
    return None
