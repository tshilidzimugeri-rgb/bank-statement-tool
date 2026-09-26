"""Consistency checks that catch misread, skipped or missing transactions,
using the running balance printed on bank statements: each balance must
equal the previous balance plus that row's credit minus its debit.
"""
from __future__ import annotations

from dataclasses import dataclass

# Amounts are parsed to cents; this only absorbs float rounding.
TOLERANCE = 0.005


@dataclass
class BalanceBreak:
    index: int  # position of the row whose balance doesn't follow
    expected: float
    actual: float


def balance_breaks(
    rows: list[tuple[float | None, float | None, float | None]], opening: float | None = None
) -> list[BalanceBreak]:
    """rows are (debit, credit, balance) in statement order. A row without a
    balance can't be checked and restarts the chain from the next balance.
    """
    breaks = []
    previous = opening
    for i, (debit, credit, balance) in enumerate(rows):
        if balance is None:
            previous = None
            continue
        if previous is not None:
            expected = previous + (credit or 0) - (debit or 0)
            if abs(expected - balance) > TOLERANCE:
                breaks.append(BalanceBreak(i, round(expected, 2), balance))
        previous = balance
    return breaks


@dataclass
class WorkbookGap:
    account: str
    after_date: object
    at_date: object
    description: str
    expected: float
    actual: float


def workbook_gaps(rows: list[dict]) -> list[WorkbookGap]:
    """Balance-chain breaks across the whole workbook (rows as read from the
    Transactions sheet, in date order), per account. A break between two
    statements means a statement for that period is missing; within one, a
    transaction was edited or misread.
    """
    by_account: dict[str, list[dict]] = {}
    for row in rows:
        account = f"{row.get('Client') or ''} / {row.get('Bank') or ''} {row.get('Account') or ''}".strip()
        by_account.setdefault(account, []).append(row)

    gaps = []
    for account, account_rows in by_account.items():
        chain = [(r.get("Debit"), r.get("Credit"), r.get("Balance")) for r in account_rows]
        for b in balance_breaks(chain):
            previous = next(
                (account_rows[i] for i in range(b.index - 1, -1, -1) if account_rows[i].get("Balance") is not None),
                account_rows[b.index],
            )
            row = account_rows[b.index]
            gaps.append(WorkbookGap(account, previous.get("Date"), row.get("Date"), row.get("Description") or "",
                                    b.expected, b.actual))
    return gaps
