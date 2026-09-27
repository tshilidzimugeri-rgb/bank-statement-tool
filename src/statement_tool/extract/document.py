"""A PDF can hold one statement or many - e.g. a year of monthly statements
scanned into one file. This splits a document into its statements and reads
each on its own, so each is checked against its own balances and totals.

Scanned (image-only) documents are read with OCR, twice and independently
(different resolution), and the two readings are compared row by row: any
row they disagree on is marked REVIEW_REQUIRED. OCR output is never
corrected - a misread figure shows up as a row that doesn't add up.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pdfplumber

from ..config import BankLayout, ClientRule, Settings
from ..models import StatementResult
from . import ocr_extract
from .bank_detect import detect_bank, extract_statement_period
from .parser import (
    PdfPasswordError,
    _find_account_number,
    _find_password,
    _period_range,
    build_result,
    parse_statement,
)
from .reading import auto_reading, compare

# A page with less extractable text than this is treated as a scanned image.
MIN_TEXT_CHARS_PER_PAGE = 40


def parse_document(
    pdf_path: Path,
    *,
    layouts: dict[str, BankLayout],
    generic: BankLayout,
    client_rules: list[ClientRule],
    settings: Settings,
    sender: str | None = None,
    subject: str | None = None,
    interactive: bool = True,
    progress=None,
) -> list[StatementResult]:
    """One StatementResult per statement in the PDF. progress(text) is called
    with status updates during slow steps (OCR)."""
    say = progress or (lambda _msg: None)
    filename = pdf_path.name
    try:
        password = _find_password(pdf_path, settings.pdf_passwords)
        with pdfplumber.open(pdf_path, password=password or "") as pdf:
            page_texts = [p.extract_text() or "" for p in pdf.pages]
    except PdfPasswordError as exc:
        return [StatementResult(source_file=filename, ok=False, transactions=[], error=str(exc),
                                status="REVIEW_REQUIRED")]
    except Exception as exc:
        return [StatementResult(source_file=filename, ok=False, transactions=[], status="REVIEW_REQUIRED",
                                error=f"Could not open PDF: {exc!r}")]

    scanned = bool(page_texts) and all(len(t.strip()) < MIN_TEXT_CHARS_PER_PAGE for t in page_texts)
    second_texts = None
    if scanned:
        say(f"Reading {len(page_texts)} scanned page(s) with OCR (first of two independent readings)...")
        try:
            page_texts = ocr_extract.ocr_pages(pdf_path, password=password, tesseract_cmd=settings.tesseract_cmd)
            say("Reading the scanned pages a second time, independently, to cross-check...")
            second_texts = ocr_extract.ocr_pages(pdf_path, password=password, tesseract_cmd=settings.tesseract_cmd,
                                                 dpi=ocr_extract.RETRY_DPI)
        except Exception as exc:
            return [StatementResult(source_file=filename, ok=False, transactions=[], used_ocr=True,
                                    status="REVIEW_REQUIRED", error=f"OCR failed: {exc}")]

    segments = split_statements(page_texts, layouts, generic)
    if len(segments) == 1 and not scanned:
        # One digital statement: the full reader, with its own second extraction.
        return [parse_statement(pdf_path, layouts=layouts, generic=generic, client_rules=client_rules,
                                settings=settings, sender=sender, subject=subject, interactive=interactive)]

    if not scanned:
        from .text_extract import second_engine_pages
        try:
            second_texts = second_engine_pages(pdf_path, password)
        except Exception:
            second_texts = None
        if second_texts is not None and len(second_texts) != len(page_texts):
            second_texts = None

    results = []
    for i, pages in enumerate(segments, 1):
        label = filename if len(segments) == 1 else f"{filename} (statement {i} of {len(segments)})"
        text = chr(10).join(page_texts[p] for p in pages)
        second = chr(10).join(second_texts[p] for p in pages) if second_texts else None
        results.append(parse_statement_text(
            text, label, second_text=second, layouts=layouts, generic=generic, client_rules=client_rules,
            sender=sender, subject=subject, interactive=interactive, used_ocr=scanned))
    return results


def split_statements(page_texts: list[str], layouts: dict[str, BankLayout], generic: BankLayout) -> list[list[int]]:
    """Page indexes of each statement: a new one starts on a page printing a
    different statement period (continuation pages print none)."""
    segments: list[list[int]] = []
    current_period = None
    for index, text in enumerate(page_texts):
        period = None
        lower = text.lower()
        if "statement period" in lower or "from date" in lower:
            period = extract_statement_period(text, detect_bank(text, layouts, generic), generic)
        if not segments or (period and period != current_period and current_period is not None):
            segments.append([])
        if period:
            current_period = period
        segments[-1].append(index)
    return segments


def parse_statement_text(
    text: str,
    label: str,
    *,
    second_text: str | None,
    layouts: dict[str, BankLayout],
    generic: BankLayout,
    client_rules: list[ClientRule],
    sender: str | None,
    subject: str | None,
    interactive: bool,
    used_ocr: bool,
) -> StatementResult:
    layout = detect_bank(text, layouts, generic)
    period = extract_statement_period(text, layout, generic)
    account = _find_account_number(text, layout, generic)
    period_range = _period_range(period)

    first = auto_reading(text, period_range, text, layout, generic)
    second = auto_reading(second_text, period_range, second_text, layout, generic) if second_text else None
    # The cleaner of the two independent readings is kept; the other checks it.
    reading, other = first, second
    if second is not None and second.rows and (not first.rows or second.score() < first.score()):
        reading, other = second, first
    if not reading.rows:
        reading = None
    else:
        compare(reading, other, "the other OCR pass" if used_ocr else "second PDF text engine")

    return build_result(reading, filename=label, full_text=text, first_page_text=text, layout=layout,
                        generic=generic, account_number=account, period=period, client_rules=client_rules,
                        sender=sender, subject=subject, interactive=interactive, used_ocr=used_ocr,
                        error_hint=("no transactions could be read from the scanned pages - the scan may be too "
                                    "unclear") if used_ocr else None)


if __name__ == "__main__":  # pragma: no cover - manual check: python -m statement_tool.extract.document file.pdf
    from .. import config as config_mod

    s = config_mod.load_settings()
    lay, gen = config_mod.load_bank_layouts(s.banks_config)
    for r in parse_document(Path(sys.argv[1]), layouts=lay, generic=gen, client_rules=[], settings=s,
                            interactive=False, progress=print):
        print(r.source_file, r.statement_period, r.status, len(r.transactions), r.problems)
