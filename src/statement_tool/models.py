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
    # Why this row needs checking by a person ("" when the statement's own
    # balances confirm it). Such rows are still added, highlighted.
    check: str = ""
    evidence: str = ""  # the line as read from the document
    confidence: float = 1.0  # 1.00 verified; below 0.95 means REVIEW_REQUIRED
    status: str = "APPROVED"  # APPROVED or REVIEW_REQUIRED
    # An amount the document doesn't show as money in or out: kept here, out
    # of the debit/credit totals, rather than guessed into one of them.
    unassigned: float | None = None


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
    # Things about the statement as a whole that a person should check (its
    # printed totals don't match, no account number, ...). The statement is
    # still added; these are listed with it.
    problems: list[str] = field(default_factory=list)
    status: str = "APPROVED"  # APPROVED only when every row is and everything reconciles
    report: dict = field(default_factory=dict)  # opening, credits, debits, net, closing (computed and printed)
