"""Reads one digital (text-based) statement: bank detection, table and
line extraction, and the checks. Scanned and multi-statement PDFs go through
document.parse_document, which uses this for the single-statement case.
"""
from __future__ import annotations

import re
import sys
from datetime import date
from pathlib import Path

import pdfplumber
from pdfminer.pdfdocument import PDFPasswordIncorrect

from ..config import BankLayout, ClientRule, Settings
from ..checks import TOLERANCE, balance_breaks
from ..models import StatementResult, Transaction
from . import line_extract, text_extract
from .amounts import parse_amount
from .auto_extract import date_in_period
from .bank_detect import detect_bank, extract_statement_period, match_client
from .dates import parse_date
from .reading import REVIEW_REQUIRED, Reading, Row, assess, auto_reading, choose, compare, finalize


class PdfPasswordError(Exception):
    pass


def _is_password_error(exc: BaseException) -> bool:
    return isinstance(exc, PDFPasswordIncorrect) or any(isinstance(a, PDFPasswordIncorrect) for a in exc.args)


def _find_password(pdf_path: Path, passwords: list[str]) -> str | None:
    """Returns None for an unencrypted PDF, else the first configured
    password that opens it; raises PdfPasswordError if none do.
    """
    for password in [None, *passwords]:
        try:
            with pdfplumber.open(pdf_path, password=password or ""):
                return password
        except Exception as exc:
            if not _is_password_error(exc):
                raise
    tried = f"none of the {len(passwords)} password(s)" if passwords else "no passwords are"
    raise PdfPasswordError(
        f"PDF is password-protected and {tried} in STATEMENT_PDF_PASSWORDS (.env) opened it"
    )


def _quick_first_page_text(pdf_path: Path, password: str | None) -> str:
    with pdfplumber.open(pdf_path, password=password or "") as pdf:
        if not pdf.pages:
            return ""
        return pdf.pages[0].extract_text() or ""


def _resolve_client(
    client_rules: list[ClientRule],
    *,
    sender: str | None,
    subject: str | None,
    filename: str,
    first_page_text: str,
    interactive: bool,
) -> tuple[str, bool]:
    """Returns (client_name, was_guessed_or_unmapped)."""
    matched = match_client(
        client_rules, sender=sender, subject=subject, filename=filename, first_page_text=first_page_text
    )
    if matched:
        return matched, False

    if interactive and sys.stdin.isatty():
        print(f"\n  Could not match a client for: {filename}")
        if sender:
            print(f"    From: {sender}")
        snippet = " ".join(first_page_text.split())[:160]
        if snippet:
            print(f"    Statement text starts: {snippet}...")
        try:
            typed = input("  Type the client name for this statement (blank to skip mapping): ").strip()
        except EOFError:
            typed = ""
        if typed:
            return typed, False

    return "UNMAPPED_CLIENT", True


def parse_statement(
    pdf_path: Path,
    *,
    layouts: dict[str, BankLayout],
    generic: BankLayout,
    client_rules: list[ClientRule],
    settings: Settings,
    sender: str | None = None,
    subject: str | None = None,
    interactive: bool = True,
) -> StatementResult:
    filename = pdf_path.name
    try:
        password = _find_password(pdf_path, settings.pdf_passwords)
        first_page_text = _quick_first_page_text(pdf_path, password)
    except PdfPasswordError as exc:
        return StatementResult(source_file=filename, ok=False, transactions=[], error=str(exc))
    except Exception as exc:  # corrupt / unreadable PDF
        return StatementResult(
            source_file=filename, ok=False, transactions=[], error=f"Could not open PDF: {exc!r}"
        )

    bank_layout = detect_bank(first_page_text, layouts, generic)
    account_number = _find_account_number(first_page_text, bank_layout, generic)

    try:
        text_result = text_extract.extract(pdf_path, bank_layout, password)
    except Exception as exc:
        text_result = None
        text_extract_error = str(exc)
    else:
        text_extract_error = None

    full_text = text_result.full_text if text_result else first_page_text
    period = extract_statement_period(full_text, bank_layout, generic)
    period_range = _period_range(period)

    # Extraction 1: the layout-independent reader first, then the column and
    # known-line readers; the one with the fewest unconfirmed rows is kept.
    readings: list[Reading] = []
    if text_result and not text_result.likely_scanned:
        readings.append(auto_reading(text_result.full_text, period_range, full_text, bank_layout, generic))
    if text_result and text_result.rows:
        table_rows = _table_rows(text_result.rows, bank_layout)
        if table_rows:
            readings.append(assess(table_rows, None, full_text, bank_layout, generic, period=period_range))
    if text_result and not text_result.likely_scanned:
        readings.extend(_line_readings(text_result.full_text, period, full_text, bank_layout, generic))
    best = choose(readings)

    # Extraction 2: the same document through an independent PDF text engine.
    if best is not None:
        try:
            second_text = text_extract.second_engine_text(pdf_path, password)
            second = auto_reading(second_text, period_range, second_text, bank_layout, generic)
        except Exception:
            second = None
        compare(best, second, "second PDF text engine")

    return build_result(best, filename=filename, full_text=full_text, first_page_text=first_page_text,
                        layout=bank_layout, generic=generic, account_number=account_number, period=period,
                        client_rules=client_rules, sender=sender, subject=subject, interactive=interactive,
                        used_ocr=False, error_hint=text_extract_error)


def _table_rows(raw_rows, layout: BankLayout) -> list[Row]:
    rows = []
    for row in raw_rows:
        debit = credit = None
        if layout.amount_style == "signed":
            amt = parse_amount(row.cells.get("amount"))
            if amt is None:
                continue
            if amt < 0:
                debit = abs(amt)
            else:
                credit = amt
        else:
            debit = parse_amount(row.cells.get("debit"))
            credit = parse_amount(row.cells.get("credit"))
            if debit is None and credit is None:
                continue
            # The column says which way the money moved; some banks also
            # print debits as negative numbers, which mustn't flip it back.
            debit = abs(debit) if debit is not None else None
            credit = abs(credit) if credit is not None else None
        balance = parse_amount(row.cells.get("balance")) if "balance" in row.cells else None
        date_raw = row.cells.get("date", "")
        description = (row.cells.get("description") or "").strip() or "(no description)"
        evidence = " ".join(v for v in row.cells.values() if v)
        rows.append(Row(parse_date(date_raw) or date_raw or None, description, debit, credit, balance,
                        evidence=evidence))
    return rows


def _line_readings(text, period, full_text, layout, generic) -> list[Reading]:
    """One reading per known line shape."""
    out = []
    for fmt in _line_formats_to_try(layout):
        result = line_extract.extract(
            text,
            thousands=fmt.get("thousands", ","),
            fee_column=bool(fmt.get("fee_column", False)),
            dates_without_year=bool(fmt.get("dates_without_year", False)),
            trailing_charges_column=bool(fmt.get("trailing_charges_column", False)),
            unmarked_is_debit=bool(fmt.get("unmarked_is_debit", False)),
        )
        rows = []
        for r in result.rows:
            date_raw = _with_year(r.date_raw, period) if fmt.get("dates_without_year") else r.date_raw
            rows.append(Row(parse_date(date_raw) or date_raw or None, r.description, r.debit, r.credit, r.balance,
                            unassigned=r.unassigned, evidence=r.evidence))
        if rows:
            out.append(assess(rows, result.opening_balance, full_text, layout, generic,
                              skipped_lines=result.skipped_lines, period=_period_range(period)))
    return out


def build_result(best: Reading | None, *, filename, full_text, first_page_text, layout, generic, account_number,
                 period, client_rules, sender, subject, interactive, used_ocr, error_hint=None) -> StatementResult:
    client, client_unmapped = _resolve_client(client_rules, sender=sender, subject=subject, filename=filename,
                                              first_page_text=first_page_text, interactive=interactive)
    base = dict(source_file=filename, bank_key=layout.key, bank_display_name=layout.display_name,
                account_number=account_number, client=client, statement_period=period or "Unknown",
                used_ocr=used_ocr)
    if best is None:
        reason = error_hint or ("no transactions could be found in it - it may not be a bank statement, "
                                "or the scan is too unclear to read")
        return StatementResult(ok=False, transactions=[], error=reason, status=REVIEW_REQUIRED, **base)

    warnings: list[str] = []
    if client_unmapped:
        warnings.append("Client could not be matched from config/clients.yaml - add a rule or re-run interactively")
    if not account_number:
        # Without it the statement can't be put with its account's other
        # statements; it gets a workbook of its own so nothing is mixed.
        best.problems.append("no account number found, so this statement was kept in a workbook of its own")

    confidences, statuses, final = finalize(best, scanned=used_ocr)
    transactions = [
        Transaction(client=client, bank=layout.display_name, statement_period=period or "Unknown",
                    date=row.date or "", description=row.description, debit=row.debit, credit=row.credit,
                    balance=row.balance, source_file=filename, ocr=used_ocr, account=account_number or "",
                    check=row.check, evidence=row.evidence, confidence=confidence, status=status,
                    unassigned=row.unassigned)
        for row, confidence, status in zip(best.rows, confidences, statuses)
    ]
    return StatementResult(ok=True, transactions=transactions, warning="; ".join(warnings) or None,
                           problems=list(best.problems), status=final, report=dict(best.report), **base)


# Line shapes seen so far (Standard Bank, Capitec, FNB, plain), tried in turn
# on any statement whose own layout doesn't read cleanly.
LINE_FORMAT_VARIANTS = [
    {},
    {"thousands": " "},
    {"thousands": " ", "fee_column": True},
    {"fee_column": True},
    {"dates_without_year": True, "trailing_charges_column": True, "unmarked_is_debit": True},
    {"dates_without_year": True, "trailing_charges_column": True},
    {"dates_without_year": True},
    {"trailing_charges_column": True},
]


def _line_formats_to_try(layout: BankLayout) -> list[dict]:
    formats = [layout.line_format or {}]
    for variant in LINE_FORMAT_VARIANTS:
        if variant not in formats:
            formats.append(variant)
    return formats


def _read_lines(text, fmt, period, full_text, layout, generic):
    """One line-by-line reading of the statement and its problems."""
    result = line_extract.extract(
        text,
        thousands=fmt.get("thousands", ","),
        fee_column=bool(fmt.get("fee_column", False)),
        dates_without_year=bool(fmt.get("dates_without_year", False)),
        trailing_charges_column=bool(fmt.get("trailing_charges_column", False)),
        unmarked_is_debit=bool(fmt.get("unmarked_is_debit", False)),
    )
    rows = []
    for r in result.rows:
        date_raw = _with_year(r.date_raw, period) if fmt.get("dates_without_year") else r.date_raw
        rows.append((parse_date(date_raw) or date_raw, r.description, r.debit, r.credit, r.balance))
    if not rows:
        return rows, []
    problems = _all_problems(rows, result.opening_balance, full_text, layout, generic)
    if result.skipped_lines:
        shown = "; ".join(result.skipped_lines[:3])
        problems.append(
            f"{len(result.skipped_lines)} line(s) look like transactions but couldn't be read (e.g. {shown})"
        )
    return rows, problems


def _period_range(period: str | None) -> tuple[date, date] | None:
    """First and last date of a statement period like "8 August 2026 to 10 September 2026"."""
    if not period:
        return None
    parts = re.split(r"\s+to\s+", period, flags=re.IGNORECASE)
    if len(parts) != 2:
        return None
    start, end = parse_date(parts[0]), parse_date(parts[1])
    if not (start and end) or start > end:
        return None
    return date.fromisoformat(start), date.fromisoformat(end)


def _period_end(period: str | None) -> date | None:
    """Last date of a statement period like "8 August 2026 to 10 September 2026"."""
    if not period:
        return None
    iso = parse_date(re.split(r"\s+to\s+", period, flags=re.IGNORECASE)[-1])
    return date.fromisoformat(iso) if iso else None


def _with_year(date_raw: str, period: str | None) -> str:
    """Adds the year to a "11 Aug" date: the one that puts it inside the
    statement period (a period can run Dec -> Jan, or over several months).
    Left as-is (and so reported as unreadable) if the period is unknown.
    """
    period_range = _period_range(period)
    period_end = period_range[1] if period_range else _period_end(period)
    if period_end is None:
        return date_raw
    iso = parse_date(f"{date_raw} 2000")  # a leap year, so 29 Feb reads too
    if iso is None:
        return date_raw
    day = date.fromisoformat(iso)
    found = date_in_period(day.month, day.day, period_range[0] if period_range else None, period_end)
    return f"{date_raw} {found.year}" if found else date_raw


def _find_account_number(first_page_text: str, layout: BankLayout, generic: BankLayout) -> str | None:
    for pattern in [*layout.account_patterns, *generic.account_patterns]:
        m = re.search(pattern, first_page_text or "", re.IGNORECASE)
        if m:
            return re.sub(r"\D", "", m.group(1))
    return None


def _all_problems(raw_transactions, opening_balance, full_text, layout, generic) -> list[str]:
    return [
        *_check_transactions(raw_transactions, opening_balance),
        *_check_control_totals(raw_transactions, full_text, layout, generic),
    ]


def _find_control_totals(name: str, text: str, layout: BankLayout, generic: BankLayout) -> list[float]:
    """Every figure the first matching pattern finds, in order. A 3- or
    6-month statement is often monthly statements put together, each month
    printing its own totals and closing balance."""
    for pattern in [*layout.control_totals.get(name, []), *generic.control_totals.get(name, [])]:
        found = [parse_amount(m.group(1)) for m in re.finditer(pattern, text, re.IGNORECASE)]
        found = [v for v in found if v is not None]
        if found:
            return found
    return []


def _find_control_total(name: str, text: str, layout: BankLayout, generic: BankLayout) -> float | None:
    found = _find_control_totals(name, text, layout, generic)
    return found[0] if found else None


def statement_total(found: list[float], read: float | None) -> float | None:
    """The total a statement prints for all of it: its one total; or, where
    each month prints its own, the months added up - or a grand total printed
    besides them (the largest), whichever the transactions read agree with."""
    if not found:
        return None
    if len(found) == 1:
        return found[0]
    together, grand = round(sum(found), 2), max(found)
    if read is not None and abs(grand - read) <= TOLERANCE:
        return grand
    return together


def printed_fees(text: str, layout: BankLayout, generic: BankLayout) -> float | None:
    """Fees the statement totals apart from its total debits, if it does."""
    found = _find_control_totals("total_fees", text, layout, generic)
    return statement_total([abs(v) for v in found], None)


def _check_control_totals(
    raw_transactions: list[tuple[str | None, str, float | None, float | None, float | None]],
    full_text: str,
    layout: BankLayout,
    generic: BankLayout,
    include_closing: bool = True,
) -> list[str]:
    """Compares what was read with totals the statement prints about
    itself - the only way to notice a missed first or last transaction.
    """
    if not raw_transactions:
        return []
    problems = []
    read_debits = sum(r[2] or 0 for r in raw_transactions)
    read_credits = sum(r[3] or 0 for r in raw_transactions)
    last_balance = next((r[4] for r in reversed(raw_transactions) if r[4] is not None), None)

    checks = [
        ("total_debits", "money out", read_debits),
        ("total_credits", "money in", read_credits),
        ("closing_balance", "closing balance", last_balance),
    ]
    for name, label, read in checks:
        if name == "closing_balance" and not include_closing:
            continue
        found = _find_control_totals(name, full_text, layout, generic)
        if not found or read is None:
            continue
        if name == "closing_balance":
            printed = found[-1]  # the last month's closing balance is the statement's
        else:
            # some statements print payments as negative
            printed = statement_total([abs(v) for v in found], read)
        if name == "total_debits" and abs(printed - read) > TOLERANCE:
            fees = printed_fees(full_text, layout, generic)
            if fees is not None:
                without_fees = statement_total([abs(v) for v in found], read - fees)
                if abs(without_fees + fees - read) <= TOLERANCE:
                    continue  # the fees are totalled apart from the debits, and together they match
        if abs(printed - read) > TOLERANCE:
            what = (f"{label} on the statement is {printed:,.2f}" if len(found) == 1 or name == "closing_balance"
                    else f"{label} printed for each month of the statement adds up to {printed:,.2f}")
            problems.append(f"{what} but the transactions read add up to {read:,.2f} - a transaction is "
                            f"missing or misread")
    return problems


def _check_transactions(
    raw_transactions: list[tuple[str | None, str, float | None, float | None, float | None]],
    opening_balance: float | None,
) -> list[str]:
    """Problems that make a statement's transactions untrustworthy."""
    if not raw_transactions:
        return []
    problems = []

    bad_dates = [r for r in raw_transactions if not parse_date(r[0])]
    if bad_dates:
        problems.append(
            f"{len(bad_dates)} transaction(s) have a date that couldn't be read "
            f"(e.g. {bad_dates[0][0]!r} on {bad_dates[0][1]!r}) - they would be left out of the monthly reports"
        )

    balances = [r[4] for r in raw_transactions]
    if all(b is None for b in balances):
        problems.append("the statement has no running balance, so the amounts can't be cross-checked")
    else:
        missing = sum(1 for b in balances if b is None)
        if missing:
            problems.append(f"{missing} transaction(s) have no balance, so they can't be cross-checked")
        breaks = balance_breaks([(r[2], r[3], r[4]) for r in raw_transactions], opening_balance)
        if breaks:
            first = raw_transactions[breaks[0].index]
            problems.append(
                f"{len(breaks)} transaction(s) don't add up with the running balance - a row may be misread "
                f"or missing (first: {first[0]} {first[1]!r}, expected balance {breaks[0].expected:,.2f}, "
                f"statement shows {breaks[0].actual:,.2f})"
            )
    return problems
