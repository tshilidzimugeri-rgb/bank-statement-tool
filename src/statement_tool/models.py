"""The normalized schema every bank layout gets mapped into."""
from __future__ import annotations

from dataclasses import dataclass


@dataclass
class Transaction:
    client: str
    bank: str
    statement_period: str
    date: str
    description: str
    debit: float | None
    credit: float | None
    balance: float | None
    source_file: str
    ocr: bool = False


@dataclass
class StatementResult:
    """Outcome of trying to parse one PDF."""

    source_file: str
    ok: bool
    transactions: list[Transaction]
    bank_key: str | None = None
    bank_display_name: str | None = None
    client: str | None = None
    statement_period: str | None = None
    used_ocr: bool = False
    error: str | None = None
    warning: str | None = None
