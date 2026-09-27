"""Text-layer extraction using pdfplumber - the primary path for statements
that were produced digitally rather than scanned.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pdfplumber

from ..config import BankLayout
from .column_map import map_header_row


@dataclass
class RawRow:
    cells: dict[str, str]
    page_number: int


@dataclass
class TextExtractionResult:
    first_page_text: str
    full_text: str
    rows: list[RawRow]
    tables_found: int
    likely_scanned: bool


# A page with fewer than this many characters of extractable text is treated
# as (probably) an image-only/scanned page for the purposes of the
# likely_scanned heuristic.
MIN_CHARS_FOR_TEXT_PAGE = 40

TABLE_SETTINGS_CANDIDATES = [
    {},  # pdfplumber defaults (lines-based)
    {
        "vertical_strategy": "text",
        "horizontal_strategy": "text",
    },
]


def extract(pdf_path: Path, layout: BankLayout, password: str | None = None) -> TextExtractionResult:
    rows: list[RawRow] = []
    tables_found = 0
    first_page_text = ""
    full_text_parts: list[str] = []
    text_page_count = 0
    total_pages = 0

    active_mapping: dict[str, int] | None = None
    active_column_count: int | None = None

    with pdfplumber.open(pdf_path, password=password or "") as pdf:
        total_pages = len(pdf.pages)
        for page_index, page in enumerate(pdf.pages):
            page_text = page.extract_text() or ""
            full_text_parts.append(page_text)
            if page_index == 0:
                first_page_text = page_text
            if len(page_text.strip()) >= MIN_CHARS_FOR_TEXT_PAGE:
                text_page_count += 1

            page_tables: list[list[list[str | None]]] = []
            for settings in TABLE_SETTINGS_CANDIDATES:
                try:
                    found = page.extract_tables(settings) if settings else page.extract_tables()
                except Exception:
                    found = []
                if found:
                    page_tables = found
                    break

            for table in page_tables:
                if not table:
                    continue
                header, *body = table
                mapping = map_header_row(header, layout)
                start_idx = 0
                if mapping is None and active_mapping is not None and len(header) == active_column_count:
                    # Likely a continuation table on a later page whose
                    # header row wasn't repeated - reuse the last mapping
                    # and treat this table's first row as data, not header.
                    mapping = active_mapping
                    body = table
                    start_idx = 0
                elif mapping is not None:
                    active_mapping = mapping
                    active_column_count = len(header)

                if mapping is None:
                    continue

                tables_found += 1
                for row in body[start_idx:]:
                    cells = {
                        col: (row[idx] if idx < len(row) else None) or ""
                        for col, idx in mapping.items()
                    }
                    if not any(v.strip() for v in cells.values()):
                        continue
                    rows.append(RawRow(cells=cells, page_number=page_index + 1))

    likely_scanned = total_pages > 0 and text_page_count == 0

    return TextExtractionResult(
        first_page_text=first_page_text,
        full_text="\n".join(full_text_parts),
        rows=rows,
        tables_found=tables_found,
        likely_scanned=likely_scanned,
    )


def second_engine_text(pdf_path: Path, password: str | None = None) -> str:
    return chr(10).join(second_engine_pages(pdf_path, password))


def second_engine_pages(pdf_path: Path, password: str | None = None) -> list[str]:
    """The document's text as read by a second, independent engine (MuPDF),
    from each word's position on the page - used to cross-check the first
    reading. Words are grouped into lines by their vertical position."""
    import pymupdf

    doc = pymupdf.open(pdf_path)
    try:
        if doc.needs_pass:
            doc.authenticate(password or "")
        pages = []
        for page in doc:
            lines: list[list] = []  # [centre_y, height, words]
            for w in sorted(page.get_text("words"), key=lambda w: ((w[1] + w[3]) / 2, w[0])):
                centre, height = (w[1] + w[3]) / 2, w[3] - w[1]
                if lines and abs(lines[-1][0] - centre) <= max(1.5, 0.4 * min(height, lines[-1][1])):
                    lines[-1][2].append(w)
                else:
                    lines.append([centre, height, [w]])
            pages.append(chr(10).join(" ".join(x[4] for x in sorted(ws, key=lambda x: x[0])) for _, _, ws in lines))
        return pages
    finally:
        doc.close()
