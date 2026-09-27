"""OCR for scanned / image-only statement PDFs.

Each page is rendered (PyMuPDF), turned the right way up (scanners often
feed pages upside down - Tesseract's orientation detection says by how
much), and read with Tesseract. The text then goes through the same line
reader and balance/total checks as a digital statement, so a misread digit
shows up as a statement that doesn't add up, never as a wrong number.
"""
from __future__ import annotations

import os
import re
import shutil
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pymupdf
import pytesseract
from PIL import Image

RENDER_DPI = 300
# Second, sharper reading for statements whose first reading doesn't check
# out - slower, but it recovers digits the first one lost.
RETRY_DPI = 400
# --psm 6: one uniform block, so each statement row stays on one line;
# preserve_interword_spaces keeps the gaps between columns.
TESSERACT_CONFIG = "--psm 6 -c preserve_interword_spaces=1"
_WINDOWS_DEFAULT = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")


class OcrUnavailableError(Exception):
    pass


def _configure_tesseract(tesseract_cmd: str | None) -> None:
    cmd = tesseract_cmd or os.environ.get("TESSERACT_CMD") or shutil.which("tesseract")
    if not cmd and _WINDOWS_DEFAULT.exists():
        cmd = str(_WINDOWS_DEFAULT)
    if not cmd:
        raise OcrUnavailableError(
            "this looks like a scanned statement, and reading scans needs Tesseract OCR, which isn't installed "
            "(Windows: winget install UB-Mannheim.TesseractOCR)"
        )
    pytesseract.pytesseract.tesseract_cmd = cmd


def _upright(image: Image.Image) -> Image.Image:
    try:
        osd = pytesseract.image_to_osd(image)
    except pytesseract.TesseractError:
        return image  # too little text to tell; read as-is
    match = re.search(r"Rotate: (\d+)", osd)
    rotate = int(match.group(1)) if match else 0
    return image.rotate(-rotate, expand=True) if rotate else image


def _ocr_page(image: Image.Image) -> str:
    return pytesseract.image_to_string(_upright(image), config=TESSERACT_CONFIG)


def ocr_pages(
    pdf_path: Path,
    *,
    password: str | None = None,
    tesseract_cmd: str | None = None,
    page_numbers: list[int] | None = None,
    dpi: int = RENDER_DPI,
) -> list[str]:
    """Text of each page (or of page_numbers, 0-based), in that order."""
    _configure_tesseract(tesseract_cmd)
    # Tesseract runs as a separate process per page, so a few pages run in
    # parallel - rendered a batch at a time, since a page image at 300 dpi
    # is ~9 MB and a long scan held all at once runs out of memory.
    workers = max(1, min(2, os.cpu_count() or 1))
    texts: list[str] = []
    doc = pymupdf.open(pdf_path)
    try:
        if doc.needs_pass and not doc.authenticate(password or ""):
            raise OcrUnavailableError("the PDF password didn't open it for OCR")
        wanted = list(range(doc.page_count)) if page_numbers is None else page_numbers
        with ThreadPoolExecutor(max_workers=workers) as pool:
            for start in range(0, len(wanted), workers):
                batch = []
                for number in wanted[start:start + workers]:
                    pix = doc[number].get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY)
                    batch.append(Image.frombytes("L", (pix.width, pix.height), pix.samples))
                texts.extend(clean_ocr_text(t) for t in pool.map(_ocr_page, batch))
    finally:
        doc.close()
    return texts


_TABLE_LINES = re.compile(r"[|\[\]{}]")
# "1,472.60Cr)" / "129.74 Cr]" - junk straight after a Cr/Dr marker.
_MARKER_JUNK = re.compile(r"(\d\.\d{2})\s?(Cr|Dr)[)\]}|!:;,.]+")


# "15.00)" - a closing bracket read off a table rule after an amount.
_AMOUNT_BRACKET = re.compile(r"(?<![(\d])(\d[\d,]*\.\d{2}(?:Cr|Dr)?)\)")
# Quote marks read after an amount: "2,299.24Cr'".
_AMOUNT_QUOTES = re.compile(r'(\d\.\d{2}(?:Cr|Dr)?)[\'`"\u2018\u2019]+')
# Junk before a row's date at the start of a line: "(02 Mar ...".
_LEADING_JUNK = re.compile(r"(?m)^[^\w\s]{1,3}\s?(?=\d{1,2} [A-Z][a-z]{2}\b)")
# Stray marks standing alone between columns: "361.79Cr * 8.00".
_LONE_MARKS = re.compile(r"(?<=\s)[^\w\s#():-]{1,3}(?=\s|$)")


def clean_ocr_text(text: str) -> str:
    """Removes what scanning adds: table rules read as | [ ] { }, stray marks
    between columns, and junk around amounts. Amounts are never changed
    (not even a decimal comma) and never guessed at - anything still misread is caught by the
    balance checks.
    """
    text = text.replace("\ufffd", " ")
    text = _MARKER_JUNK.sub(r"\1\2", text)
    text = _TABLE_LINES.sub(" ", text)
    text = _AMOUNT_BRACKET.sub(r"\1", text)
    text = _AMOUNT_QUOTES.sub(r"\1", text)
    text = _LEADING_JUNK.sub("", text)
    text = _LONE_MARKS.sub(" ", text)
    return "\n".join(re.sub(r"[ \t]+", " ", line).strip() for line in text.splitlines())
