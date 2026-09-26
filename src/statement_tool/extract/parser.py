"""Orchestrates bank detection, text/OCR extraction and normalisation into
the common Transaction schema for a single PDF.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pdfplumber
from pdfminer.pdfdocument import PDFPasswordIncorrect

from ..config import BankLayout, ClientRule, Settings
from ..checks import TOLERANCE, balance_breaks
from ..models import StatementResult, Transaction
from . import line_extract, ocr_extract, text_extract
from .amounts import parse_amount
from .bank_detect import detect_bank, extract_statement_period, match_client
from .dates import parse_date


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

    used_ocr = False
    ocr_error: str | None = None
    full_text = text_result.full_text if text_result else first_page_text
    warnings: list[str] = []

    raw_transactions: list[tuple[str | None, str, float | None, float | None, float | None]] = []
    # (date_raw, description, debit, credit, balance)

    if text_result and text_result.rows:
        for row in text_result.rows:
            debit = credit = None
            if bank_layout.amount_style == "signed":
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
            date_iso = parse_date(date_raw)
            description = (row.cells.get("description") or "").strip() or "(no description)"
            raw_transactions.append((date_iso or date_raw, description, debit, credit, balance))

    problems: list[str] = []
    if raw_transactions:
        problems = _all_problems(raw_transactions, None, full_text, bank_layout, generic)
    if (not raw_transactions or problems) and text_result and not text_result.likely_scanned:
        # No usable table, or the table's numbers don't check out: read the
        # transaction lines directly, and keep whichever reading checks out.
        line_result = line_extract.extract(
            text_result.full_text,
            thousands=bank_layout.line_format.get("thousands", ","),
            fee_column=bool(bank_layout.line_format.get("fee_column", False)),
        )
        line_rows = [
            (parse_date(r.date_raw) or r.date_raw, r.description, r.debit, r.credit, r.balance)
            for r in line_result.rows
        ]
        if line_rows:
            line_problems = _all_problems(line_rows, line_result.opening_balance, full_text, bank_layout, generic)
            if line_result.skipped_lines:
                shown = "; ".join(line_result.skipped_lines[:3])
                line_problems.append(
                    f"{len(line_result.skipped_lines)} line(s) look like transactions but couldn't be read "
                    f"(e.g. {shown})"
                )
            if not raw_transactions or not line_problems:
                raw_transactions, problems = line_rows, line_problems

    if not raw_transactions:
        # Text extraction found no usable transactions - fall back to OCR.
        used_ocr = True
        try:
            ocr_result = ocr_extract.extract(
                pdf_path,
                tesseract_cmd=settings.tesseract_cmd,
                poppler_path=settings.poppler_path,
                password=password,
            )
        except Exception as exc:
            ocr_error = str(exc)
            ocr_result = None

        if ocr_result:
            full_text = ocr_result.full_text
            for row in ocr_result.rows:
                debit = credit = None
                if row.amount is not None:
                    if row.is_credit_guess:
                        credit = row.amount if row.amount >= 0 else abs(row.amount)
                    else:
                        debit = abs(row.amount)
                date_iso = parse_date(row.date_raw) or row.date_raw
                raw_transactions.append((date_iso, row.description, debit, credit, row.balance))
            if raw_transactions:
                warnings.append(
                    "Extracted via OCR - column alignment is approximate; "
                    "please spot-check debit/credit amounts against the PDF"
                )

    if used_ocr and raw_transactions:
        problems = _all_problems(raw_transactions, None, full_text, bank_layout, generic)

    statement_period = extract_statement_period(full_text, bank_layout, generic) or "Unknown"
    client, client_unmapped = _resolve_client(
        client_rules,
        sender=sender,
        subject=subject,
        filename=filename,
        first_page_text=first_page_text,
        interactive=interactive,
    )
    if client_unmapped:
        warnings.append("Client could not be matched from config/clients.yaml - add a rule or re-run interactively")
    if not account_number and raw_transactions:
        # Without it the statement can't be matched to its account's
        # workbook, and statements must never be mixed.
        problems.append(
            "no account number found on the statement, so it can't be put with its account's other "
            "statements - add an account_patterns entry for this bank in config/banks.yaml "
            "(if added anyway, it gets a workbook of its own)"
        )

    if not raw_transactions:
        error_bits = []
        if text_extract_error:
            error_bits.append(f"text extraction error: {text_extract_error}")
        if ocr_error:
            error_bits.append(f"OCR error: {ocr_error}")
        if not error_bits:
            error_bits.append(
                "no transaction rows recognized via text extraction or OCR - "
                "this bank's layout may need a new entry in config/banks.yaml"
            )
        return StatementResult(
            source_file=filename,
            ok=False,
            transactions=[],
            bank_key=bank_layout.key,
            bank_display_name=bank_layout.display_name,
            client=client,
            statement_period=statement_period,
            used_ocr=used_ocr,
            error="; ".join(error_bits),
        )

    transactions = [
        Transaction(
            client=client,
            bank=bank_layout.display_name,
            statement_period=statement_period,
            date=date_val or "",
            description=description,
            debit=debit,
            credit=credit,
            balance=balance,
            source_file=filename,
            ocr=used_ocr,
            account=account_number or "",
        )
        for date_val, description, debit, credit, balance in raw_transactions
    ]

    return StatementResult(
        source_file=filename,
        ok=True,
        transactions=transactions,
        bank_key=bank_layout.key,
        bank_display_name=bank_layout.display_name,
        account_number=account_number,
        client=client,
        statement_period=statement_period,
        used_ocr=used_ocr,
        warning="; ".join(warnings) if warnings else None,
        problems=problems,
    )


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


def _find_control_total(name: str, text: str, layout: BankLayout, generic: BankLayout) -> float | None:
    for pattern in [*layout.control_totals.get(name, []), *generic.control_totals.get(name, [])]:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            return parse_amount(m.group(1))
    return None


def _check_control_totals(
    raw_transactions: list[tuple[str | None, str, float | None, float | None, float | None]],
    full_text: str,
    layout: BankLayout,
    generic: BankLayout,
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
        printed = _find_control_total(name, full_text, layout, generic)
        if printed is None or read is None:
            continue
        if name != "closing_balance":
            printed = abs(printed)  # some statements print payments as negative
        if abs(printed - read) > TOLERANCE:
            problems.append(
                f"{label} on the statement is {printed:,.2f} but the transactions read add up to "
                f"{read:,.2f} - a transaction is missing or misread"
            )
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
