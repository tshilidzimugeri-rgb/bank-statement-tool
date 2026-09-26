"""A PDF can hold one statement or many - e.g. a year of monthly statements
scanned into one file. This splits a document into its statements and reads
each on its own, so each is checked against its own opening/closing balance
and printed totals.

Scanned (image-only) documents are read with OCR. OCR misreads the odd digit,
so for scans a row that doesn't add up can be repaired from the running
balance - but a repair is kept only if the whole statement then matches its
printed totals and closing balance exactly, which a wrong repair can't do.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pdfplumber

from ..checks import TOLERANCE
from ..config import BankLayout, ClientRule, Settings
from ..models import StatementResult, Transaction
from . import line_extract, ocr_extract
from .bank_detect import detect_bank, extract_statement_period
from .dates import parse_date
from .parser import (
    PdfPasswordError,
    _all_problems,
    _find_account_number,
    _find_control_total,
    _find_password,
    _line_formats_to_try,
    _read_lines,
    _resolve_client,
    parse_statement,
)

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
        return [StatementResult(source_file=filename, ok=False, transactions=[], error=str(exc))]
    except Exception as exc:
        return [StatementResult(source_file=filename, ok=False, transactions=[], error=f"Could not open PDF: {exc!r}")]

    scanned = bool(page_texts) and all(len(t.strip()) < MIN_TEXT_CHARS_PER_PAGE for t in page_texts)
    if scanned:
        say(f"Reading {len(page_texts)} scanned page(s) with OCR - this takes a few seconds per page...")
        try:
            page_texts = ocr_extract.ocr_pages(pdf_path, password=password, tesseract_cmd=settings.tesseract_cmd)
        except Exception as exc:
            return [StatementResult(source_file=filename, ok=False, transactions=[], used_ocr=True,
                                    error=f"OCR failed: {exc}")]

    segments = split_statements(page_texts, layouts, generic)
    if len(segments) == 1 and not scanned:
        # One digital statement: the full reader (tables, then lines).
        return [parse_statement(pdf_path, layouts=layouts, generic=generic, client_rules=client_rules,
                                settings=settings, sender=sender, subject=subject, interactive=interactive)]

    def read(pages: list[int], texts: list[str], label: str) -> StatementResult:
        return parse_statement_text(chr(10).join(texts[p] for p in pages), label, layouts=layouts, generic=generic,
                                    client_rules=client_rules, sender=sender, subject=subject,
                                    interactive=interactive, used_ocr=scanned)

    results = []
    for i, pages in enumerate(segments, 1):
        label = filename if len(segments) == 1 else f"{filename} (statement {i} of {len(segments)})"
        result = read(pages, page_texts, label)
        if scanned and (not result.ok or result.problems):
            say(f"{label}: re-reading {len(pages)} page(s) more sharply...")
            try:
                sharper = ocr_extract.ocr_pages(pdf_path, password=password, tesseract_cmd=settings.tesseract_cmd,
                                                page_numbers=pages, dpi=ocr_extract.RETRY_DPI)
            except Exception:
                sharper = None
            if sharper is not None:
                retry_texts = list(page_texts)
                for number, text in zip(pages, sharper):
                    retry_texts[number] = text
                retried = read(pages, retry_texts, label)
                if retried.ok and not retried.problems:
                    result = retried
        results.append(result)
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

    best = None
    for fmt in _line_formats_to_try(layout):
        rows, problems = _read_lines(text, fmt, period, text, layout, generic)
        if rows and (best is None or len(problems) < len(best[1])):
            best = (rows, problems, fmt)
        if rows and not problems:
            break

    warnings: list[str] = []
    rows, problems = ([], []) if best is None else (best[0], best[1])
    if problems and used_ocr:
        opening = _opening_balance(text, best[2])
        repaired, notes = repair_from_balances(rows, opening, text, layout, generic)
        if repaired is not None:
            rows, problems = repaired, []
            shown = "; ".join(notes[:3]) + ("; ..." if len(notes) > 3 else "")
            warnings.append(
                f"scanned statement: {len(notes)} misread figure(s) corrected from the running balance and "
                f"confirmed by the statement's own totals ({shown})"
            )

    client, unmapped = _resolve_client(client_rules, sender=sender, subject=subject, filename=label,
                                       first_page_text=text, interactive=interactive)
    if unmapped:
        warnings.append("Client could not be matched from config/clients.yaml - add a rule or re-run interactively")
    if rows and not account:
        problems.append(
            "no account number found on the statement, so it can't be put with its account's other statements"
        )

    base = dict(source_file=label, bank_key=layout.key, bank_display_name=layout.display_name,
                account_number=account, client=client, statement_period=period or "Unknown", used_ocr=used_ocr)
    if not rows:
        what = "the scanned pages" if used_ocr else "the statement"
        return StatementResult(ok=False, transactions=[], error=f"no transaction rows recognized in {what}", **base)

    transactions = [
        Transaction(client=client, bank=layout.display_name, statement_period=period or "Unknown",
                    date=d or "", description=desc, debit=debit, credit=credit, balance=balance,
                    source_file=label, ocr=used_ocr, account=account or "")
        for d, desc, debit, credit, balance in rows
    ]
    return StatementResult(ok=True, transactions=transactions, warning="; ".join(warnings) or None,
                           problems=problems, **base)


def _opening_balance(text: str, fmt: dict) -> float | None:
    return line_extract.extract(
        text,
        thousands=fmt.get("thousands", ","),
        fee_column=bool(fmt.get("fee_column", False)),
        dates_without_year=bool(fmt.get("dates_without_year", False)),
        trailing_charges_column=bool(fmt.get("trailing_charges_column", False)),
        unmarked_is_debit=bool(fmt.get("unmarked_is_debit", False)),
    ).opening_balance


def _close(a: float | None, b: float | None) -> bool:
    return a is not None and b is not None and abs(a - b) <= TOLERANCE


def repair_from_balances(rows, opening, text, layout, generic):
    """Fixes single misread figures using the running balance:

    - a misread balance: the row's amount and the next row both agree on
      what the balance must have been;
    - a misread amount: the balance before and after it are consistent with
      the rest of the chain, so the amount is their difference;
    - an unreadable date ("41 Mar"): only when the rows either side share
      one date, so there's no other possibility.

    Returns (repaired rows, notes) only if the repaired statement then passes
    every check including the printed totals and closing balance; otherwise
    (None, []).
    """
    printed = [_find_control_total(n, text, layout, generic) for n in ("total_debits", "total_credits")]
    closing = _find_control_total("closing_balance", text, layout, generic)
    if all(p is None for p in printed) or closing is None or opening is None:
        return None, []  # nothing independent to confirm a repair with

    rows = [list(r) for r in rows]
    notes: list[str] = []

    def balance_before(i: int) -> float | None:
        return opening if i == 0 else rows[i - 1][4]

    def follows(i: int, balance: float) -> bool:
        prev = balance_before(i)
        return prev is not None and _close(prev + (rows[i][3] or 0) - (rows[i][2] or 0), balance)

    def next_follows(i: int, balance: float) -> bool:
        if i + 1 < len(rows):
            nxt = rows[i + 1]
            return _close(balance + (nxt[3] or 0) - (nxt[2] or 0), nxt[4])
        return _close(balance, closing)

    for i in range(len(rows)):
        balance = rows[i][4]
        if balance is None or follows(i, balance):
            continue
        prev = balance_before(i)
        if prev is None:
            return None, []
        date, desc, debit, credit = rows[i][:4]
        expected = round(prev + (credit or 0) - (debit or 0), 2)
        if next_follows(i, expected):
            rows[i][4] = expected
            notes.append(f"{date} {desc[:30]}: balance read as {balance:,.2f}, is {expected:,.2f}")
        elif next_follows(i, balance):
            delta = round(balance - prev, 2)
            old = debit if debit is not None else credit
            rows[i][2], rows[i][3] = (None, delta) if delta > 0 else (-delta, None)
            notes.append(f"{date} {desc[:30]}: amount read as {old or 0:,.2f}, is {abs(delta):,.2f}")
        else:
            return None, []  # more than one figure wrong here - can't be sure

    for i, row in enumerate(rows):
        if parse_date(row[0]) is None:
            fixed = _only_possible_date(rows, i)
            if fixed is None:
                return None, []
            notes.append(f"{row[1][:30]}: date read as {row[0]!r}, is {fixed}")
            row[0] = fixed

    repaired = [tuple(r) for r in rows]
    if _all_problems(repaired, opening, text, layout, generic):
        return None, []
    return repaired, notes


def _only_possible_date(rows: list[list], i: int) -> str | None:
    """For an unreadable date between two readable rows on the same date,
    that date - the only one it can be."""
    before = next((parse_date(rows[j][0]) for j in range(i - 1, -1, -1) if parse_date(rows[j][0])), None)
    after = next((parse_date(rows[j][0]) for j in range(i + 1, len(rows)) if parse_date(rows[j][0])), None)
    return before if before and before == after else None


if __name__ == "__main__":  # pragma: no cover - manual check: python -m statement_tool.extract.document file.pdf
    from .. import config as config_mod

    s = config_mod.load_settings()
    lay, gen = config_mod.load_bank_layouts(s.banks_config)
    for r in parse_document(Path(sys.argv[1]), layouts=lay, generic=gen, client_rules=[], settings=s,
                            interactive=False, progress=print):
        print(r.source_file, r.statement_period, len(r.transactions), r.problems or "OK", r.warning or "")