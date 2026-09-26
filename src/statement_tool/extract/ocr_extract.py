"""OCR fallback for scanned / image-only bank statement PDFs.

This is deliberately best-effort: OCR loses the column boundaries that the
text-based extractor relies on, so we fall back to a per-line heuristic
(leading date, trailing amount(s), last amount on the line = balance) rather
than a real table structure. Every statement that goes through this path is
flagged (StatementResult.used_ocr) so it always surfaces in the run summary
for a manual spot-check - the numbers should be treated as provisional.
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass
from pathlib import Path

import pytesseract
from pdf2image import convert_from_path

from .amounts import parse_amount
from .dates import parse_date

_DATE_TOKEN = re.compile(
    r"^\s*(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}|\d{1,2}\s+[A-Za-z]{3,9}\.?\s+\d{2,4})\s*"
)
_AMOUNT_TOKEN = re.compile(
    r"(?<![A-Za-z0-9])(?:R\s?)?-?\(?\d{1,3}(?:[ ,]\d{3})*(?:[.,]\d{2})\)?\s?(?:Cr|Dr)?(?![A-Za-z0-9])",
    re.IGNORECASE,
)


@dataclass
class OcrRow:
    date_raw: str
    description: str
    amount: float | None
    is_credit_guess: bool | None
    balance: float | None


@dataclass
class OcrExtractionResult:
    full_text: str
    rows: list[OcrRow]


def _render_and_ocr(pdf_path: Path, poppler_path: str | None, password: str | None) -> str:
    kwargs = {}
    if poppler_path:
        kwargs["poppler_path"] = poppler_path
    if password:
        kwargs["userpw"] = password
    images = convert_from_path(str(pdf_path), dpi=300, **kwargs)
    texts = []
    for image in images:
        texts.append(pytesseract.image_to_string(image))
    return "\n".join(texts)


def extract(
    pdf_path: Path, *, tesseract_cmd: str | None, poppler_path: str | None, password: str | None = None
) -> OcrExtractionResult:
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    elif "TESSERACT_CMD" in os.environ:
        pytesseract.pytesseract.tesseract_cmd = os.environ["TESSERACT_CMD"]

    full_text = _render_and_ocr(pdf_path, poppler_path, password)

    rows: list[OcrRow] = []
    for line in full_text.splitlines():
        date_match = _DATE_TOKEN.match(line)
        if not date_match:
            continue
        date_raw = date_match.group(1)
        if parse_date(date_raw) is None:
            continue

        remainder = line[date_match.end():]
        amount_matches = list(_AMOUNT_TOKEN.finditer(remainder))
        if not amount_matches:
            continue

        description = remainder[: amount_matches[0].start()].strip(" -|:\t")

        balance = parse_amount(amount_matches[-1].group())
        amount = None
        is_credit_guess = None
        if len(amount_matches) >= 2:
            amount_text = amount_matches[-2].group()
            amount = parse_amount(amount_text)
            if amount is not None:
                is_credit_guess = "dr" not in amount_text.lower() and amount >= 0

        rows.append(
            OcrRow(
                date_raw=date_raw,
                description=description or "(description not confidently read)",
                amount=amount,
                is_credit_guess=is_credit_guess,
                balance=balance,
            )
        )

    return OcrExtractionResult(full_text=full_text, rows=rows)
