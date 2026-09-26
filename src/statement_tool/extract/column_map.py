"""Map a table's header row onto the logical columns a bank layout expects."""
from __future__ import annotations

import re

from ..config import BankLayout


def _clean(cell: str | None) -> str:
    return re.sub(r"\s+", " ", (cell or "")).strip().lower()


def map_header_row(header_row: list[str | None], layout: BankLayout) -> dict[str, int] | None:
    """Return {logical_column: cell_index} if this header row looks like the
    transaction table's header for the given layout, else None.

    A row qualifies if it contains at least a date-like column, a
    description-like column, and either (debit & credit) or (amount) -
    whichever the layout's amount_style calls for - plus ideally balance.
    """
    cleaned = [_clean(c) for c in header_row]

    def find(aliases: list[str]) -> int | None:
        for idx, cell in enumerate(cleaned):
            if not cell:
                continue
            if any(alias in cell for alias in aliases):
                return idx
        return None

    date_idx = find(layout.columns.get("date", ["date"]))
    desc_idx = find(layout.columns.get("description", ["description"]))
    if date_idx is None or desc_idx is None:
        return None

    mapping = {"date": date_idx, "description": desc_idx}

    if layout.amount_style == "signed":
        amount_idx = find(layout.columns.get("amount", ["amount"]))
        if amount_idx is None:
            return None
        mapping["amount"] = amount_idx
    else:
        debit_idx = find(layout.columns.get("debit", ["debit"]))
        credit_idx = find(layout.columns.get("credit", ["credit"]))
        if debit_idx is None and credit_idx is None:
            return None
        if debit_idx is not None:
            mapping["debit"] = debit_idx
        if credit_idx is not None:
            mapping["credit"] = credit_idx

    balance_idx = find(layout.columns.get("balance", ["balance"]))
    if balance_idx is not None:
        mapping["balance"] = balance_idx

    # Two headings in one cell (e.g. "Debit Credit" merged by the table
    # finder) means the columns can't be told apart - reading it anyway would
    # put every amount in both. Reject it so the line-based reader is used.
    if len(set(mapping.values())) < len(mapping):
        return None

    return mapping
