"""Appends normalized transactions into one combined Excel workbook.

Design choices, since they matter for correctness:

- Everything lives on a single "Transactions" sheet (per the brief - no
  per-client tabs) formatted as a real Excel Table so AutoFilter/sorting and
  PivotTables work immediately.
- A hidden "Row Key" column (sha256 of the row's normalized fields) gives a
  second, transaction-level dedup guard on top of the statement-level
  ProcessedStore - belt and braces against ever double-counting a row.
- A TOTALS row below the table uses real =SUM(...) formulas over the
  current data range, rebuilt every run, so it recalculates if you edit a
  cell by hand rather than being a snapshot number.
- A "Summary" sheet is rebuilt from scratch each run with SUMIF/COUNTIF
  formulas per client, so it always reflects the live Transactions data.
"""
from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from .models import Transaction

TRANSACTIONS_SHEET = "Transactions"
SUMMARY_SHEET = "Summary"
TABLE_NAME = "TransactionsTable"

HEADERS = [
    "Client",
    "Bank",
    "Statement Period",
    "Date",
    "Description",
    "Debit",
    "Credit",
    "Balance",
    "Source File",
    "OCR",
    "Row Key",
]
CURRENCY_FORMAT = '"R" #,##0.00'
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


@dataclass
class WriteResult:
    added: int
    skipped_duplicates: int


def _row_key(t: Transaction) -> str:
    raw = "|".join(
        [
            t.client,
            t.bank,
            t.statement_period,
            t.date,
            t.description,
            f"{t.debit:.2f}" if t.debit is not None else "",
            f"{t.credit:.2f}" if t.credit is not None else "",
            f"{t.balance:.2f}" if t.balance is not None else "",
            t.source_file,
        ]
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]


def _get_or_create_transactions_sheet(wb: Workbook) -> Worksheet:
    if TRANSACTIONS_SHEET in wb.sheetnames:
        return wb[TRANSACTIONS_SHEET]
    if wb.sheetnames == ["Sheet"] and wb["Sheet"].max_row == 1 and wb["Sheet"].max_column == 1 and wb["Sheet"]["A1"].value is None:
        ws = wb["Sheet"]
        ws.title = TRANSACTIONS_SHEET
    else:
        ws = wb.create_sheet(TRANSACTIONS_SHEET)
    # Written via explicit cell() calls, not ws.append(): a brand-new
    # Workbook()'s default sheet starts with _current_row already at 1
    # (an openpyxl quirk), which would push append() output to row 2.
    for col_idx, header in enumerate(HEADERS, start=1):
        ws.cell(row=1, column=col_idx, value=header)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    widths = [22, 16, 24, 12, 42, 13, 13, 13, 28, 8, 4]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.column_dimensions[get_column_letter(len(HEADERS))].hidden = True
    return ws


def _existing_data_rows(ws: Worksheet) -> tuple[list[list], int]:
    """Returns (data_rows_as_values, first_row_index) reading rows 2.. until
    a row with a blank Row Key is hit (that's a totals/separator row from a
    previous run, which we always regenerate) or the sheet ends.
    """
    data_rows = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        row_key_cell = row[len(HEADERS) - 1]
        if row_key_cell.value in (None, ""):
            break
        data_rows.append([c.value for c in row])
    return data_rows, len(data_rows)


def _clear_below_header(ws: Worksheet) -> None:
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)


def _write_date_cell(ws: Worksheet, row_idx: int, col_idx: int, value: str) -> None:
    cell = ws.cell(row=row_idx, column=col_idx)
    if DATE_RE.match(value or ""):
        cell.value = datetime.strptime(value, "%Y-%m-%d").date()
        cell.number_format = "yyyy-mm-dd"
    else:
        cell.value = value or ""


def _rebuild_table(ws: Worksheet, last_row: int) -> None:
    if TABLE_NAME in ws.tables:
        del ws.tables[TABLE_NAME]
    ref = f"A1:{get_column_letter(len(HEADERS))}{last_row}"
    table = Table(displayName=TABLE_NAME, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    ws.add_table(table)


def append_transactions(workbook_path: Path, transactions: list[Transaction]) -> WriteResult:
    workbook_path.parent.mkdir(parents=True, exist_ok=True)
    wb = load_workbook(workbook_path) if workbook_path.exists() else Workbook()
    ws = _get_or_create_transactions_sheet(wb)

    existing_rows, data_row_count = _existing_data_rows(ws)
    existing_keys = {row[len(HEADERS) - 1] for row in existing_rows}

    _clear_below_header(ws)
    for row_idx, row_values in enumerate(existing_rows, start=2):
        for col_idx, value in enumerate(row_values, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    added = 0
    skipped = 0
    next_row = data_row_count + 2
    for t in transactions:
        key = _row_key(t)
        if key in existing_keys:
            skipped += 1
            continue
        existing_keys.add(key)
        ws.cell(row=next_row, column=1, value=t.client)
        ws.cell(row=next_row, column=2, value=t.bank)
        ws.cell(row=next_row, column=3, value=t.statement_period)
        _write_date_cell(ws, next_row, 4, t.date)
        ws.cell(row=next_row, column=5, value=t.description)
        debit_cell = ws.cell(row=next_row, column=6, value=t.debit)
        debit_cell.number_format = CURRENCY_FORMAT
        credit_cell = ws.cell(row=next_row, column=7, value=t.credit)
        credit_cell.number_format = CURRENCY_FORMAT
        balance_cell = ws.cell(row=next_row, column=8, value=t.balance)
        balance_cell.number_format = CURRENCY_FORMAT
        ws.cell(row=next_row, column=9, value=t.source_file)
        ws.cell(row=next_row, column=10, value="Yes" if t.ocr else "No")
        ws.cell(row=next_row, column=11, value=key)
        next_row += 1
        added += 1

    last_data_row = next_row - 1
    if last_data_row >= 2:
        totals_row = last_data_row + 2
        ws.cell(row=totals_row, column=1, value="TOTAL").font = Font(bold=True)
        debit_total = ws.cell(row=totals_row, column=6, value=f"=SUM(F2:F{last_data_row})")
        debit_total.font = Font(bold=True)
        debit_total.number_format = CURRENCY_FORMAT
        credit_total = ws.cell(row=totals_row, column=7, value=f"=SUM(G2:G{last_data_row})")
        credit_total.font = Font(bold=True)
        credit_total.number_format = CURRENCY_FORMAT
        count_cell = ws.cell(row=totals_row, column=9, value=f"=COUNTA(A2:A{last_data_row})")
        count_cell.font = Font(italic=True)
        _rebuild_table(ws, last_data_row)

    _rebuild_summary_sheet(wb, last_data_row if last_data_row >= 2 else 0)
    wb.save(workbook_path)
    return WriteResult(added=added, skipped_duplicates=skipped)


def _rebuild_summary_sheet(wb: Workbook, last_data_row: int) -> None:
    if SUMMARY_SHEET in wb.sheetnames:
        ws = wb[SUMMARY_SHEET]
        wb.remove(ws)
    ws = wb.create_sheet(SUMMARY_SHEET)
    headers = ["Client", "Bank", "Transaction Count", "Total Debit", "Total Credit", "Net"]
    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=header)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for idx, width in enumerate([26, 18, 16, 15, 15, 15], start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    if last_data_row < 2:
        wb.move_sheet(SUMMARY_SHEET, offset=-(len(wb.sheetnames) - 1))
        return

    tx_ws = wb[TRANSACTIONS_SHEET]
    pairs: set[tuple[str, str]] = set()
    for row in tx_ws.iter_rows(min_row=2, max_row=last_data_row, values_only=True):
        client, bank = row[0], row[1]
        if client:
            pairs.add((client, bank or ""))

    row_idx = 2
    for client, bank in sorted(pairs):
        ws.cell(row=row_idx, column=1, value=client)
        ws.cell(row=row_idx, column=2, value=bank)
        rng_client = f"Transactions!$A$2:$A${last_data_row}"
        rng_bank = f"Transactions!$B$2:$B${last_data_row}"
        rng_debit = f"Transactions!$F$2:$F${last_data_row}"
        rng_credit = f"Transactions!$G$2:$G${last_data_row}"
        ws.cell(
            row=row_idx, column=3,
            value=f'=COUNTIFS({rng_client},A{row_idx},{rng_bank},B{row_idx})',
        )
        debit_cell = ws.cell(
            row=row_idx, column=4,
            value=f'=SUMIFS({rng_debit},{rng_client},A{row_idx},{rng_bank},B{row_idx})',
        )
        debit_cell.number_format = CURRENCY_FORMAT
        credit_cell = ws.cell(
            row=row_idx, column=5,
            value=f'=SUMIFS({rng_credit},{rng_client},A{row_idx},{rng_bank},B{row_idx})',
        )
        credit_cell.number_format = CURRENCY_FORMAT
        net_cell = ws.cell(row=row_idx, column=6, value=f"=E{row_idx}-D{row_idx}")
        net_cell.number_format = CURRENCY_FORMAT
        row_idx += 1

    total_row = row_idx + 1
    ws.cell(row=total_row, column=1, value="GRAND TOTAL").font = Font(bold=True)
    gt_debit = ws.cell(row=total_row, column=4, value=f"=SUM(D2:D{row_idx - 1})")
    gt_debit.font = Font(bold=True)
    gt_debit.number_format = CURRENCY_FORMAT
    gt_credit = ws.cell(row=total_row, column=5, value=f"=SUM(E2:E{row_idx - 1})")
    gt_credit.font = Font(bold=True)
    gt_credit.number_format = CURRENCY_FORMAT
    gt_net = ws.cell(row=total_row, column=6, value=f"=E{total_row}-D{total_row}")
    gt_net.font = Font(bold=True)
    gt_net.number_format = CURRENCY_FORMAT

    wb.move_sheet(SUMMARY_SHEET, offset=-(len(wb.sheetnames) - 1))
