"""The normalized schema every bank layout gets mapped into."""
from __future__ import annotations

from dataclasses import dataclass, field


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
    category: str = ""
    account: str = ""


@dataclass
class StatementResult:
    """Outcome of trying to parse one PDF."""

    source_file: str
    ok: bool
    transactions: list[Transaction]
    bank_key: str | None = None
    bank_display_name: str | None = None
    account_number: str | None = None
    client: str | None = None
    statement_period: str | None = None
    used_ocr: bool = False
    error: str | None = None
    warning: str | None = None
    # Checks that failed (rows that don't add up with the running balance,
    # lines that couldn't be read, ...). Such a statement isn't trusted: it's
    # only written to the workbook if the user explicitly overrides.
    problems: list[str] = field(default_factory=list)
