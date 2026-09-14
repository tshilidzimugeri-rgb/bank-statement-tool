"""Orchestrates bank detection, text/OCR extraction and normalisation into
the common Transaction schema for a single PDF.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pdfplumber

from ..config import BankLayout, ClientRule, Settings
from ..models import StatementResult, Transaction
from . import ocr_extract, text_extract
from .amounts import parse_amount
from .bank_detect import detect_bank, extract_statement_period, match_client
from .dates import parse_date


def _quick_first_page_text(pdf_path: Path) -> str:
    with pdfplumber.open(pdf_path) as pdf:
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
        first_page_text = _quick_first_page_text(pdf_path)
    except Exception as exc:  # corrupt / unreadable PDF
        return StatementResult(source_file=filename, ok=False, transactions=[], error=f"Could not open PDF: {exc}")

    bank_layout = detect_bank(first_page_text, layouts, generic)

    try:
        text_result = text_extract.extract(pdf_path, bank_layout)
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
        unparsed_dates = 0
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

            balance = parse_amount(row.cells.get("balance")) if "balance" in row.cells else None
            date_raw = row.cells.get("date", "")
            date_iso = parse_date(date_raw)
            if date_iso is None:
                unparsed_dates += 1
            description = (row.cells.get("description") or "").strip() or "(no description)"
            raw_transactions.append((date_iso or date_raw, description, debit, credit, balance))

        if unparsed_dates:
            warnings.append(f"{unparsed_dates} row(s) had a date that could not be normalized to YYYY-MM-DD")

    else:
        # Text extraction found no usable transaction table - fall back to OCR.
        used_ocr = True
        try:
            ocr_result = ocr_extract.extract(
                pdf_path, tesseract_cmd=settings.tesseract_cmd, poppler_path=settings.poppler_path
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
        )
        for date_val, description, debit, credit, balance in raw_transactions
    ]

    return StatementResult(
        source_file=filename,
        ok=True,
        transactions=transactions,
        bank_key=bank_layout.key,
        bank_display_name=bank_layout.display_name,
        client=client,
        statement_period=statement_period,
        used_ocr=used_ocr,
        warning="; ".join(warnings) if warnings else None,
    )
