"""Line-based reader for statements whose text layer is clean but whose
table grid pdfplumber can't split into columns (e.g. Standard Bank's
6-month statement, where each whole row lands in the first cell, and
Capitec's Transaction History).

Recognizes lines shaped like:

    28 Feb 26 CARTRACK S105778598 -1,097.68 19,696.84
    DEBIT TRANSFER                      <- optional short continuation line

i.e. date, description, one amount, then the running balance - plus, for
layouts with fee_column (Capitec), an optional fee before the balance:

    03/08/2026 Banking App External PayShap Payment ... -3 205.00 -2.00 468.91

The running balance is used to confirm whether each amount is a debit or a
credit, so this also works for statements that print amounts unsigned.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from .amounts import parse_amount

_DATE = r"\d{1,2}\s+[A-Za-z]{3,9}\.?\s+\d{2,4}|\d{1,2}/\d{1,2}/\d{2,4}|\d{4}-\d{2}-\d{2}"

# A continuation line longer than this is treated as page furniture (legal
# footer, disclaimer) rather than part of the transaction's description.
MAX_CONTINUATION_CHARS = 60

# Balance arithmetic tolerance, to absorb float rounding.
_CENT = 0.005


def _amount_pattern(thousands: str) -> str:
    # Only the layout's own thousands separator is accepted, so reference
    # numbers at the end of a description are never glued onto an amount.
    sep = re.escape(thousands)
    return rf"\(?-?\d{{1,3}}(?:{sep}\d{{3}})*\.\d{{2}}\)?(?:\s?(?:Cr|Dr))?"


@dataclass
class _Patterns:
    txn: re.Pattern
    looks_like_txn: re.Pattern
    opening: re.Pattern


def _patterns(thousands: str, fee_column: bool) -> _Patterns:
    amount = _amount_pattern(thousands)
    # A trailing "*" marks a VAT-inclusive amount (Capitec); not part of the number.
    fee = rf"(?:\s+(?P<fee>{amount})\*?)?" if fee_column else ""
    return _Patterns(
        txn=re.compile(
            rf"^(?P<date>{_DATE})\s+(?P<desc>.+?)\s+(?P<amount>{amount})\*?{fee}\s+(?P<balance>{amount})$"
        ),
        # Starts with a date and has a bare amount on it (not an "R58.00"
        # summary figure) but didn't match txn - probably a transaction in a
        # shape we don't read; reported, never silently dropped.
        looks_like_txn=re.compile(rf"^(?:{_DATE})\s.*(?<![\dRr.,]){amount}"),
        opening=re.compile(rf"opening balance:?\s*-?R?\s?(?P<balance>{amount})", re.IGNORECASE),
    )


@dataclass
class LineRow:
    date_raw: str
    description: str
    debit: float | None
    credit: float | None
    balance: float | None


@dataclass
class LineExtractionResult:
    rows: list[LineRow]
    opening_balance: float | None
    skipped_lines: list[str]
    # Rows whose amount didn't reconcile with the running balance either way;
    # their debit/credit side was taken from the amount's sign instead.
    unreconciled: int


def _is_continuation(line: str, txn: re.Pattern) -> bool:
    return (
        0 < len(line) <= MAX_CONTINUATION_CHARS
        and not txn.match(line)
        and "balance" not in line.lower()
        and not line.startswith("*")  # "* Includes VAT at 15%" footnotes
    )


def extract(full_text: str, thousands: str = ",", fee_column: bool = False) -> LineExtractionResult:
    p = _patterns(thousands, fee_column)
    rows: list[LineRow] = []
    skipped: list[str] = []
    opening_balance: float | None = None
    unreconciled = 0
    prev_balance: float | None = None
    lines = [ln.strip() for ln in (full_text or "").splitlines()]

    for i, line in enumerate(lines):
        opening = p.opening.search(line)
        if opening and not rows and opening_balance is None:
            prev_balance = opening_balance = parse_amount(opening.group("balance"))
            continue

        m = p.txn.match(line)
        if not m:
            # Only count lines after the first transaction: summaries printed
            # above the transaction list (scheduled payments etc.) aren't
            # transactions, and a missed first row is caught by the balance
            # and control-total checks instead.
            if rows and p.looks_like_txn.match(line):
                skipped.append(line)
            continue

        amount = parse_amount(m.group("amount"))
        balance = parse_amount(m.group("balance"))
        fee = parse_amount(m.groupdict().get("fee"))
        if amount is None:
            continue

        description = m.group("desc").strip()
        if i + 1 < len(lines) and _is_continuation(lines[i + 1], p.txn) and lines[i + 1] != description:
            description = f"{description} | {lines[i + 1]}"

        # With a fee on the row, the printed balance is after both; the
        # balance between them is worked out so each row still chains.
        fee_amount = abs(fee) if fee is not None else 0.0
        balance_before_fee = balance + fee_amount if balance is not None else None

        magnitude = abs(amount)
        is_debit = amount < 0
        if prev_balance is not None and balance_before_fee is not None:
            if abs(prev_balance - magnitude - balance_before_fee) < _CENT:
                is_debit = True
            elif abs(prev_balance + magnitude - balance_before_fee) < _CENT:
                is_debit = False
            else:
                unreconciled += 1

        rows.append(
            LineRow(
                date_raw=m.group("date"),
                description=description,
                debit=magnitude if is_debit else None,
                credit=None if is_debit else magnitude,
                balance=round(balance_before_fee, 2) if balance_before_fee is not None else None,
            )
        )
        if fee is not None:
            rows.append(
                LineRow(
                    date_raw=m.group("date"),
                    description=f"Fee: {description}",
                    debit=fee_amount,
                    credit=None,
                    balance=balance,
                )
            )
        if balance is not None:
            prev_balance = balance

    return LineExtractionResult(
        rows=rows, opening_balance=opening_balance, skipped_lines=skipped, unreconciled=unreconciled
    )
