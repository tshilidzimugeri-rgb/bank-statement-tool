"""Checks a statement's rows and compares independent readings of it.

Rules (see README "How statements are checked"):

- Printed values are never changed, guessed or filled in. A value that
  can't be established from the document is left empty and the row is
  marked REVIEW_REQUIRED with the reason.
- Each row's balance must follow from the previous balance and its amount;
  the statement's own printed totals and closing balance must match.
- The document is read twice, independently (two PDF text engines, or two
  OCR passes for scans); any row the two readings disagree on is marked
  REVIEW_REQUIRED.
- Duplicate rows, impossible or out-of-order dates, and rows without a
  date or amount are flagged.
- A row is APPROVED (confidence 1.00; 0.97 from a scan) only when the
  balances confirm it and both readings agree. A statement is APPROVED only
  when every row is and everything reconciles.
"""
from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field
from datetime import date

from ..checks import TOLERANCE, balance_breaks
from ..config import BankLayout
from . import auto_extract
from .dates import parse_date

APPROVED = "APPROVED"
REVIEW_REQUIRED = "REVIEW_REQUIRED"
CONFIDENCE_VERIFIED = 1.00
CONFIDENCE_VERIFIED_SCAN = 0.97
CONFIDENCE_UNCONFIRMED = 0.80


@dataclass
class Row:
    date: str | None
    description: str
    debit: float | None
    credit: float | None
    balance: float | None  # as printed; None where none is printed
    unassigned: float | None = None  # amount the document doesn't show as in or out
    evidence: str = ""  # the line as read from the document
    check: str = ""  # why it needs review; "" when confirmed

    def figures(self) -> tuple:
        return (self.date, self.debit, self.credit, self.balance, self.unassigned)


@dataclass
class Reading:
    rows: list[Row]
    problems: list[str] = field(default_factory=list)  # about the statement as a whole
    report: dict = field(default_factory=dict)  # reconciliation figures

    def score(self) -> tuple:
        return (len(self.problems), sum(1 for r in self.rows if r.check), -len(self.rows))


def assess(rows: list[Row], opening: float | None, full_text: str, layout: BankLayout, generic: BankLayout,
           skipped_lines: list[str] | None = None, groups_confirmed: bool = False,
           period: tuple[date, date] | None = None) -> Reading:
    """groups_confirmed: rows printed without a balance were already
    confirmed by the next balance (the layout-independent reader does that)."""
    from .parser import (  # parser imports this module
        _check_control_totals, _find_control_total, _find_control_totals, statement_total,
    )

    def flag(row: Row, reason: str) -> None:
        if reason not in row.check:
            row.check = f"{row.check}; {reason}" if row.check else reason

    for row in rows:
        iso = parse_date(row.date) if row.date else None
        if iso:
            row.date = iso
        else:
            if row.date:
                flag(row, f"date unreadable ({row.date})")
            elif "date" not in row.check:
                flag(row, "no date")
            row.date = None
        if row.debit is None and row.credit is None and row.unassigned is None:
            flag(row, "no amount")
        elif row.unassigned is not None:
            flag(row, "the statement doesn't establish whether this money came in or went out")

    problems: list[str] = []
    printed = {n: _find_control_total(n, full_text, layout, generic)
               for n in ("total_debits", "total_credits", "closing_balance")}
    total_problems = _check_control_totals([_as_tuple(r) for r in rows], full_text, layout, generic,
                                           include_closing=False)
    closings = auto_extract.printed_closings(full_text)

    balances = [r.balance for r in rows]
    if rows and all(b is None for b in balances):
        if total_problems or all(v is None for v in printed.values()):
            problems.append("the statement prints no running balance, so its amounts couldn't be confirmed")
            for row in rows:
                flag(row, "not confirmed - the statement has no running balance")
    else:
        if not groups_confirmed:
            for row in rows:
                if row.balance is None:
                    flag(row, "no balance printed on this row to confirm it")
        for b in balance_breaks([(r.debit, r.credit, r.balance) for r in rows], opening):
            flag(rows[b.index], "does not follow from the previous balance - a row may be missing or misread")

    for row in rows:
        garbled = _garbled_figures(row.evidence)
        if garbled:
            flag(row, f"a figure contains letters ({', '.join(garbled)}) - possible OCR misread")
    _flag_duplicates(rows, flag)
    _flag_date_order(rows, flag)
    if period:
        start, end = period
        for row in rows:
            if row.date and not (start.isoformat() <= row.date <= end.isoformat()):
                flag(row, f"date outside the statement period ({start:%d %b %Y} to {end:%d %b %Y})")

    if skipped_lines:
        shown = "; ".join(s[:60] for s in skipped_lines[:3])
        problems.append(f"{len(skipped_lines)} line(s) look like transactions but couldn't be read (e.g. {shown})")
    problems.extend(total_problems)
    if rows and not closings and all(v is None for v in printed.values()):
        problems.append("no closing balance or statement totals are printed, so rows missing at the end "
                        "(or a missing last page) can't be ruled out")
    problems.extend(_missing_pages(full_text))
    last_printed = next((r.balance for r in reversed(rows) if r.balance is not None), None)
    if closings and last_printed is not None and not any(abs(c - last_printed) <= TOLERANCE for c in closings):
        problems.append(f"the statement's closing balance is {closings[0]:,.2f} but the last balance read is "
                        f"{last_printed:,.2f} - a transaction may be missing or misread")

    credits = round(sum(r.credit or 0 for r in rows), 2)
    debits = round(sum(r.debit or 0 for r in rows), 2)
    last_balance = next((r.balance for r in reversed(rows) if r.balance is not None), None)
    computed_closing = round(opening + credits - debits, 2) if opening is not None else None
    if computed_closing is not None and last_balance is not None and not any(r.unassigned for r in rows):
        if abs(computed_closing - last_balance) > TOLERANCE:
            problems.append(f"opening {opening:,.2f} + credits {credits:,.2f} - debits {debits:,.2f} = "
                            f"{computed_closing:,.2f}, but the last balance is {last_balance:,.2f}")
    if total_problems or any("closing balance is" in p for p in problems):
        # The statement's own totals prove something in it is wrong, so no
        # row can count as confirmed.
        for row in rows:
            flag(row, "the statement's printed totals or closing balance don't reconcile")
    report = {
        "opening_balance": opening, "total_credits": credits, "total_debits": debits,
        "net_movement": round(credits - debits, 2), "computed_closing": computed_closing,
        "printed_closing": next((c for c in closings if last_balance is not None
                                 and abs(c - last_balance) <= TOLERANCE), closings[0] if closings else None),
        # For the whole statement: a statement of several months can print totals for each month.
        "printed_total_credits": statement_total(
            [abs(v) for v in _find_control_totals("total_credits", full_text, layout, generic)], credits),
        "printed_total_debits": statement_total(
            [abs(v) for v in _find_control_totals("total_debits", full_text, layout, generic)], debits),
        "last_balance": last_balance,
    }
    return Reading(rows, problems, report)


# A figure with letters mixed into its digits ("12,O35.24", "1I0.00", "S00.00").
_FIGURE_WITH_LETTERS = re.compile(r"(?<![A-Za-z0-9.,])R?([0-9OIlSB][0-9OIlSB,]*[.][0-9OIlSB]{2})(?:Cr|Dr|CR|DR)?(?![A-Za-z0-9.,])")


def _garbled_figures(text: str) -> list[str]:
    return [m.group(1) for m in _FIGURE_WITH_LETTERS.finditer(text or "")
            if any(ch.isdigit() for ch in m.group(1)) and any(ch.isalpha() for ch in m.group(1))]


_PAGE_OF = re.compile(r"\b(?:page|pg)\.?\s*(\d{1,3})\s*(?:of|/)\s*(\d{1,3})\b", re.IGNORECASE)


def _missing_pages(full_text: str) -> list[str]:
    """Where pages are numbered "Page 2 of 4", every page must be there."""
    found = [(int(k), int(n)) for k, n in _PAGE_OF.findall(full_text or "") if 0 < int(k) <= int(n)]
    if not found:
        return []
    total = max(n for _, n in found)
    missing = sorted(set(range(1, total + 1)) - {k for k, _ in found})
    if not missing:
        return []
    return [f"page(s) {', '.join(map(str, missing))} of {total} are missing from the document"]


def _as_tuple(r: Row) -> tuple:
    return (r.date, r.description, r.debit, r.credit, r.balance)


def _flag_duplicates(rows: list[Row], flag) -> None:
    """Same date, description and amount twice, where the running balance
    doesn't show them as two separate transactions."""
    seen: dict = {}
    for row in rows:
        key = (row.date, row.description, row.debit, row.credit, row.unassigned)
        if key in seen:
            other = seen[key]
            if row.balance is None or other.balance is None or abs(row.balance - other.balance) <= TOLERANCE:
                flag(row, "possible duplicate of an earlier row")
        seen[key] = row


def _flag_date_order(rows: list[Row], flag) -> None:
    dates = [r.date for r in rows if r.date]
    if len(dates) < 2:
        return
    ascending = sum(a <= b for a, b in zip(dates, dates[1:])) >= sum(a >= b for a, b in zip(dates, dates[1:]))
    previous = None
    for row in rows:
        if not row.date:
            continue
        if previous and ((row.date < previous) if ascending else (row.date > previous)):
            flag(row, "date out of order with the rows around it")
        previous = row.date


def auto_reading(text: str, period: tuple[date, date] | None, full_text: str, layout: BankLayout,
                 generic: BankLayout) -> Reading:
    result = auto_extract.extract(text, period[1] if period else None, period[0] if period else None)
    rows = [Row(r.date_iso, r.description, r.debit, r.credit, r.balance, r.unassigned, r.evidence, r.check)
            for r in result.rows]
    return assess(rows, result.opening_balance, full_text, layout, generic, groups_confirmed=True, period=period)


def choose(readings: list[Reading]) -> Reading | None:
    """The reading with the fewest problems and unconfirmed rows (the first on ties)."""
    readings = [r for r in readings if r.rows]
    return min(readings, key=Reading.score) if readings else None


def compare(primary: Reading, second: Reading | None, what: str) -> None:
    """Marks rows the second, independent reading doesn't agree with."""
    if second is None or not second.rows:
        primary.problems.append(f"the second independent reading ({what}) found no transactions to compare")
        for row in primary.rows:
            if not row.check:
                row.check = f"not confirmed by a second independent reading ({what})"
        return
    a = [r.figures() for r in primary.rows]
    b = [r.figures() for r in second.rows]
    matcher = difflib.SequenceMatcher(a=a, b=b, autojunk=False)
    matched = set()
    for block in matcher.get_matching_blocks():
        matched.update(range(block.a, block.a + block.size))
    for i, row in enumerate(primary.rows):
        if i not in matched:
            note = f"the second independent reading ({what}) reads this row differently"
            row.check = f"{row.check}; {note}" if row.check else note
    extra = len(b) - sum(block.size for block in matcher.get_matching_blocks())
    if extra > 0:
        primary.problems.append(f"the second independent reading ({what}) has {extra} row(s) the first doesn't")


def finalize(reading: Reading, scanned: bool) -> tuple[list[float], list[str], str]:
    """Per-row confidence and status, and the statement's final status."""
    confidences, statuses = [], []
    for row in reading.rows:
        if row.check:
            confidences.append(CONFIDENCE_UNCONFIRMED)
            statuses.append(REVIEW_REQUIRED)
        else:
            confidences.append(CONFIDENCE_VERIFIED_SCAN if scanned else CONFIDENCE_VERIFIED)
            statuses.append(APPROVED)
    final = APPROVED if not reading.problems and all(s == APPROVED for s in statuses) else REVIEW_REQUIRED
    return confidences, statuses, final
