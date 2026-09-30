"""Writes normalized transactions into one workbook, plus financial reports
built on them.

Design choices, since they matter for correctness:

- All transactions live on one "Transactions" sheet, formatted as a real
  Excel Table (AutoFilter/sorting/PivotTables work immediately) and kept in
  date order however statements are uploaded.
- A hidden "Row Key" column (hash of the transaction's own fields - not
  which file it came from) dedups at transaction level, so overlapping
  statements (e.g. a 6-month statement plus the monthly ones covering the
  same dates) never double-count.
- Each transaction gets a Category (config/categories.yaml). A category
  typed over by hand in the sheet is kept on every later run; uncategorised
  rows are re-checked against the rules each run. What the rules still
  can't place gets a category learnt from the workbook's categorised rows
  (learn.py), when similar ones clearly agree. "Category Source" says which
  of these set each category; a suggestion is re-made each run until a
  person types over it.
- The reports (Monthly Summary, Income Statement, Cash Flow, Category
  Breakdown) are rebuilt every run with SUMIFS/COUNTIFS formulas over the
  Transactions sheet, so editing a transaction or its category recalculates
  them. Opening/closing balances are the only stored values: they come from
  the statements' printed running balance, and each report has a check row
  comparing its totals against them, which exposes missing statements or
  misread rows.
- One sheet per month ("Jul 2026") lists that month's income and expenses
  separately with 15% VAT split out, and "VAT Summary" totals them per
  month. Whether a transaction includes VAT is its "VAT" (Yes/No) cell on
  the Transactions sheet - defaulted from its category, and kept like a
  hand-edited category.
"""
from __future__ import annotations

import calendar
import hashlib
from collections import Counter
import os
import re
import shutil
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from .categorize import UNCATEGORISED_EXPENSE, UNCATEGORISED_INCOME, Category, categorize, with_fallbacks
from .learn import CategoryLearner
from .models import Transaction
from .xl_values import add_cached_values

TRANSACTIONS_SHEET = "Transactions"
MONTHLY_SHEET = "Monthly Summary"
INCOME_SHEET = "Income Statement"
CASHFLOW_SHEET = "Cash Flow"
BREAKDOWN_SHEET = "Category Breakdown"
CATEGORIES_SHEET = "Categories"
REVIEW_SHEET = "Review"
REVIEW_REQUIRED = "REVIEW_REQUIRED"
REVIEW_FILL = PatternFill("solid", fgColor="FCE4D6")
APPROVED_FILL = PatternFill("solid", fgColor="E2EFDA")
VAT_SUMMARY_SHEET = "VAT Summary"
# Per-month VAT sheets are named like "Jul 2026"; all are rebuilt each run.
MONTH_SHEET_RE = re.compile(r"^[A-Z][a-z]{2} \d{4}$")
VAT_RATE = 0.15
# Sheets rebuilt from scratch every run ("Summary" is the pre-reports
# per-client sheet, dropped if an older workbook still has it).
GENERATED_SHEETS = (
    MONTHLY_SHEET, INCOME_SHEET, CASHFLOW_SHEET, BREAKDOWN_SHEET, VAT_SUMMARY_SHEET, CATEGORIES_SHEET, "Summary",
    REVIEW_SHEET, "VAT201",
)
# Month sheets go between the reports and Transactions, in date order.
REPORT_SHEET_ORDER = (REVIEW_SHEET, MONTHLY_SHEET, INCOME_SHEET, CASHFLOW_SHEET, BREAKDOWN_SHEET, VAT_SUMMARY_SHEET,
                      "VAT201")
TRAILING_SHEET_ORDER = (TRANSACTIONS_SHEET, CATEGORIES_SHEET)
TABLE_NAME = "TransactionsTable"

HEADERS = [
    "Client",
    "Bank",
    "Account",
    "Statement Period",
    "Date",
    "Month",
    "Description",
    "Category",
    "Category Source",  # Rule, Set by you, or Suggested (learnt from similar rows)
    "VAT",
    "Debit",
    "Credit",
    "In/Out Not Shown",  # amounts the statement doesn't show as money in or out - in no total
    "Balance",
    "Status",  # APPROVED or REVIEW_REQUIRED
    "Confidence",
    "Review Reason",
    "Evidence",  # the line as read from the document
    "Source File",
    "OCR",
    "Row Key",
    # Hidden: the category last suggested, so a person typing over it is told
    # apart from the suggestion itself.
    "Suggested Category",
]
WIDTHS = [18, 14, 13, 22, 12, 9, 48, 30, 26, 6, 13, 13, 14, 14, 17, 11, 50, 70, 28, 6, 4, 4]
CATEGORY_BY_RULE = "Rule"
CATEGORY_BY_HAND = "Set by you"
COL = {header: get_column_letter(i) for i, header in enumerate(HEADERS, start=1)}

CURRENCY_FORMAT = '"R" #,##0.00;[Red]-"R" #,##0.00'
PERCENT_FORMAT = "0.0%"
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
TYPE_ORDER = ("income", "expense", "transfer", "drawings")

BOLD = Font(bold=True)
TITLE = Font(bold=True, size=14)
HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
TOTAL_FILL = PatternFill("solid", fgColor="C6E0B4")
TOTAL_BORDER = Border(top=Side(style="thin"))


# Copies of the workbook as it was before each write, newest kept.
BACKUPS_KEPT = 30


class MixedAccountsError(Exception):
    """Transactions from a different bank account than the workbook's.
    Each account has its own workbook, so its balances and totals stay
    about one account only."""


class WorkbookLockedError(Exception):
    """The workbook couldn't be replaced - almost always because it's open
    in Excel. Nothing was changed."""


@dataclass
class WriteResult:
    added: int
    skipped_duplicates: int


def _row_key(t: Transaction, occurrence: int = 0) -> str:
    """occurrence numbers repeats of an identical row within one statement:
    two genuinely separate same-day, same-amount transactions can even share
    a balance (e.g. two R2.00 fees with money coming in between), and would
    otherwise look like duplicates, dropping one. The same statement period
    numbers its repeats the same way, so overlapping statements still match.
    """
    return _key(t.client, t.bank, t.account, t.date, t.description, t.debit, t.credit, t.balance, t.unassigned,
                occurrence)


def _key(client, bank, account, day, description, debit, credit, balance, unassigned, occurrence) -> str:
    raw = "|".join(
        [
            client or "",
            bank or "",
            *([str(account)] if account else []),
            day or "",
            _same_words(description),
            f"{debit:.2f}" if debit is not None else "",
            f"{credit:.2f}" if credit is not None else "",
            f"{balance:.2f}" if balance is not None else "",
            *([f"u{unassigned:.2f}"] if unassigned is not None else []),
            *([str(occurrence)] if occurrence else []),
        ]
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]


def _same_words(description: str | None) -> str:
    """The description's words in a fixed order. Two statements covering the
    same days can wrap a long description differently ("... Fees | Za" in
    one, "... Za Fees" in the other); with the same date, amounts and
    balance it's the same transaction. Only the matching uses this - the
    description is kept as printed."""
    return " ".join(sorted(re.findall(r"[^\s|]+", (description or "").lower())))


def _rekey(rows: list[dict]) -> None:
    """Row keys of a workbook's rows, worked out afresh from their own
    figures, so workbooks made by an earlier version match new statements."""
    seen: Counter = Counter()
    for row in rows:
        day = _as_date(row.get("Date"))
        day = day.isoformat() if isinstance(day, date) else str(day or "")
        fields = (row.get("Client"), row.get("Bank"), row.get("Account"), day, row.get("Description"),
                  row.get("Debit"), row.get("Credit"), row.get("Balance"), row.get("In/Out Not Shown"))
        base = _key(*fields, 0)
        row["Row Key"] = _key(*fields, seen[base])
        seen[base] += 1


def _as_date(value) -> date | str:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "")
    if DATE_RE.match(text):
        return datetime.strptime(text, "%Y-%m-%d").date()
    return text


def _month_of(value: date | str) -> str:
    return value.strftime("%Y-%m") if isinstance(value, date) else ""


def _read_existing_rows(wb: Workbook) -> list[dict]:
    """Rows 2.. of the Transactions sheet as {header: value}, stopping at the
    first row without a Row Key (the TOTAL row, which is regenerated).
    Columns are found by header name so older layouts still load.
    """
    if TRANSACTIONS_SHEET not in wb.sheetnames:
        return []
    ws = wb[TRANSACTIONS_SHEET]
    headers = [c.value for c in ws[1]]
    if "Row Key" not in headers:
        return []
    key_idx = headers.index("Row Key")
    rows = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        if values[key_idx] in (None, ""):
            break
        rows.append({h: v for h, v in zip(headers, values) if h})
    return rows


_UNSAFE_FILENAME = re.compile(r'[<>:"/\\|?*]')


def workbook_path_for(output_dir: Path, client: str | None, account: str | None, source_file: str) -> Path:
    """One workbook per bank account, e.g. "FTW Properties (10237421516).xlsx".
    A statement whose account number is unknown gets a workbook of its own,
    so it can never be mixed with another account's statements.
    """
    name = client if client and client != "UNMAPPED_CLIENT" else ""
    if account:
        label = f"{name} ({account})" if name else f"Account {account}"
    else:
        label = f"{name or 'Unknown account'} - {Path(source_file).stem}"
    return output_dir / f"{_UNSAFE_FILENAME.sub('_', label).strip()}.xlsx"


def restore_workbook(data: bytes, output_dir: Path, name: str) -> Path | None:
    """Puts a workbook downloaded earlier back under its account's own name -
    read from its rows, not the file name, since browsers rename repeat
    downloads ("... (1).xlsx"). Returns None if it isn't one of ours.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    tmp = output_dir / f"~restore-{hashlib.sha256(data).hexdigest()[:12]}.xlsx"
    tmp.write_bytes(data)
    try:
        rows = read_transactions(tmp)
    except Exception:
        rows = []
    if not rows or "Account" not in rows[0]:
        tmp.unlink(missing_ok=True)
        return None
    first = rows[0]
    target = workbook_path_for(output_dir, first.get("Client"), str(first.get("Account") or ""),
                               str(first.get("Source File") or name))
    os.replace(tmp, target)
    return target


def list_workbooks(output_dir: Path) -> list[Path]:
    """Account workbooks in output_dir, most recently updated first."""
    if not output_dir.exists():
        return []
    books = [p for p in output_dir.glob("*.xlsx") if not p.name.startswith("~")]
    return sorted(books, key=lambda p: p.stat().st_mtime, reverse=True)


def read_transactions(workbook_path: Path) -> list[dict]:
    """The workbook's transaction rows as {header: value} (empty if none yet)."""
    if not workbook_path.exists():
        return []
    wb = _open_workbook(workbook_path, read_only=True)
    try:
        return _read_existing_rows(wb)
    finally:
        # Read-only mode keeps the file open until closed; on Windows an open
        # file can't be replaced, so a later save would fail.
        wb.close()


def _open_workbook(workbook_path: Path, read_only: bool = False) -> Workbook:
    try:
        return load_workbook(workbook_path, read_only=read_only)
    except PermissionError as exc:
        raise WorkbookLockedError(
            f"{workbook_path.name} can't be opened - close it in Excel and try again. Nothing was changed."
        ) from exc


def _vat_default(category: str, categories: list[Category]) -> str:
    for c in with_fallbacks(categories):
        if c.name == category:
            return "Yes" if c.vat else "No"
    return "No"


def _uncategorised(category) -> bool:
    return category in (None, "", UNCATEGORISED_INCOME, UNCATEGORISED_EXPENSE)


def _row_from_transaction(t: Transaction, key: str, categories: list[Category]) -> dict:
    when = _as_date(t.date)
    category = t.category or categorize(t.description, bool(t.credit), categories)
    return {
        "Client": t.client,
        "Bank": t.bank,
        "Account": t.account,
        "Statement Period": t.statement_period,
        "Date": when,
        "Month": _month_of(when),
        "Description": t.description,
        "Category": category,
        "Category Source": None if _uncategorised(category) else CATEGORY_BY_HAND if t.category else CATEGORY_BY_RULE,
        "VAT": _vat_default(category, categories),
        "Debit": t.debit,
        "Credit": t.credit,
        "In/Out Not Shown": t.unassigned,
        "Balance": t.balance,
        "Status": t.status,
        "Confidence": t.confidence,
        "Review Reason": t.check,
        "Evidence": t.evidence,
        "Source File": t.source_file,
        "OCR": "Yes" if t.ocr else "No",
        "Row Key": key,
    }


def append_transactions(
    workbook_path: Path,
    transactions: list[Transaction],
    categories: list[Category] | None = None,
    statement: dict | None = None,
) -> WriteResult:
    """statement: the statement's reconciliation report for the Review sheet
    (see statement_review_row)."""
    categories = categories or []
    workbook_path.parent.mkdir(parents=True, exist_ok=True)
    wb = _open_workbook(workbook_path) if workbook_path.exists() else Workbook()

    rows = _read_existing_rows(wb)
    existing_accounts = {str(r.get("Account") or "") for r in rows}
    new_accounts = {t.account or "" for t in transactions}
    if rows and "" in new_accounts:
        # No account number to compare: only the same statement may be re-added.
        existing_sources = {str(r.get("Source File") or "") for r in rows}
        new_sources = {t.source_file for t in transactions if not t.account}
        if new_sources - existing_sources:
            raise MixedAccountsError(
                f"{workbook_path.name} holds other statements; refusing to add "
                f"{', '.join(sorted(new_sources - existing_sources))}, whose account number is unknown"
            )
    if rows and new_accounts - existing_accounts - {""}:
        raise MixedAccountsError(
            f"{workbook_path.name} holds account {', '.join(sorted(existing_accounts)) or '(none)'}; "
            f"refusing to add account {', '.join(sorted(new_accounts - existing_accounts)) or '(none)'} to it"
        )
    _rekey(rows)
    was = {}  # each row's category as it was, to tell which changed
    for row in rows:
        row["Date"] = _as_date(row.get("Date"))
        row["Month"] = _month_of(row["Date"])
        was[row["Row Key"]] = old = row.get("Category")
        by_rules = categorize(row.get("Description") or "", bool(row.get("Credit")), categories)
        if row.get("Suggested Category") and old == row["Suggested Category"]:
            row["Category"] = None  # still the suggestion: rules and learning get a fresh go
        # Uncategorised rows get another go against the current rules, so a
        # rule added later applies to them; any other category is left alone.
        if _uncategorised(row.get("Category")):
            row["Category"] = by_rules
            row["Category Source"] = None if _uncategorised(by_rules) else CATEGORY_BY_RULE
        else:
            row["Category Source"] = CATEGORY_BY_RULE if row["Category"] == by_rules else CATEGORY_BY_HAND
        row["Suggested Category"] = None
    keys = {row["Row Key"] for row in rows}

    added = skipped = 0
    seen_in_batch: dict[str, int] = {}
    for t in transactions:
        key = _row_key(t)
        occurrence = seen_in_batch.get(key, 0)
        seen_in_batch[key] = occurrence + 1
        key = _row_key(t, occurrence)
        if key in keys:
            skipped += 1
            continue
        keys.add(key)
        rows.append(_row_from_transaction(t, key, categories))
        was[key] = rows[-1]["Category"]
        added += 1

    _suggest_categories(rows)
    for row in rows:
        if row["Category"] != was[row["Row Key"]]:
            row["VAT"] = None  # newly categorised: take the new category's default
        if row.get("VAT") not in ("Yes", "No"):
            row["VAT"] = _vat_default(row["Category"], categories)

    # Stable sort: same-day rows keep their statement order.
    rows.sort(key=lambda r: (0, r["Date"]) if isinstance(r["Date"], date) else (1, date.max))

    reviews = _read_reviews(wb)
    if statement:
        reviews = [r for r in reviews if r.get("Source File") != statement.get("Source File")] + [statement]

    for name in list(wb.sheetnames):
        if name in (TRANSACTIONS_SHEET, *GENERATED_SHEETS) or MONTH_SHEET_RE.match(name):
            wb.remove(wb[name])
    for leftover in list(wb.sheetnames):  # a brand-new Workbook's blank "Sheet"
        if wb[leftover].max_row == 1 and wb[leftover]["A1"].value is None:
            wb.remove(wb[leftover])

    _write_transactions_sheet(wb, rows)
    _write_review_sheet(wb, reviews, len(rows))
    if rows:
        reported = report_categories(categories, rows)
        months = sorted({r["Month"] for r in rows if r["Month"]})
        balances = _month_balances(rows, months)
        income_rows = _write_income_statement(wb, rows, reported, months)
        _write_monthly_summary(wb, rows, months, balances)
        _write_cash_flow(wb, months, balances, income_rows)
        _write_category_breakdown(wb, rows, reported)
        month_totals = _write_month_vat_sheets(wb, rows, reported, months)
        _write_vat_summary(wb, month_totals)
        _write_vat201(wb, month_totals)
        _write_categories_sheet(wb, reported)

    month_sheets = [n for n in wb.sheetnames if MONTH_SHEET_RE.match(n)]  # created in date order
    order = [*REPORT_SHEET_ORDER, *month_sheets, *TRAILING_SHEET_ORDER]
    for position, name in enumerate(n for n in order if n in wb.sheetnames):
        wb.move_sheet(name, offset=position - wb.sheetnames.index(name))
    wb.active = 0
    _save_safely(wb, workbook_path)
    return WriteResult(added=added, skipped_duplicates=skipped)


def _suggest_categories(rows: list[dict]) -> None:
    """Rows the rules leave uncategorised get the category of similar rows
    a person categorised, when those clearly agree (learn.py). Rule-made
    categories aren't learnt from: the rules already cover what they match,
    and spreading them to near misses goes wrong ("Fee: Payment X" is a
    bank charge; "Payment X" isn't)."""
    taught = [r.get("Category Source") == CATEGORY_BY_HAND for r in rows]
    learner = CategoryLearner(
        [(r.get("Description") or "", bool(r.get("Credit")), r["Category"]) for r, t in zip(rows, taught) if t],
        [r.get("Description") or "" for r, t in zip(rows, taught) if not t])
    for row in rows:
        if not _uncategorised(row["Category"]):
            continue
        suggestion = learner.suggest(row.get("Description"), bool(row.get("Credit")))
        if suggestion:
            row["Category"] = row["Suggested Category"] = suggestion.category
            row["Category Source"] = suggestion.note()


def _save_safely(wb: Workbook, workbook_path: Path) -> None:
    """Back up the current workbook, then write the new one to a temporary
    file and swap it in, so a crash mid-save can never leave a half-written
    workbook behind.
    """
    tmp_path = workbook_path.with_name(f"~{workbook_path.stem}.saving.xlsx")
    wb.save(tmp_path)
    try:
        # Each formula's result saved beside it, so viewers that don't
        # calculate (phone and e-mail previews) show the numbers too.
        add_cached_values(tmp_path)
    except Exception:  # never lose a save over it: Excel calculates the formulas itself
        pass
    try:
        if workbook_path.exists():
            backups = workbook_path.parent / "backups"
            backups.mkdir(exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
            shutil.copy2(workbook_path, backups / f"{workbook_path.stem}-{stamp}.xlsx")
            for old in sorted(backups.glob(f"{workbook_path.stem}-*.xlsx"))[:-BACKUPS_KEPT]:
                old.unlink()
        os.replace(tmp_path, workbook_path)
    except PermissionError as exc:
        tmp_path.unlink(missing_ok=True)
        raise WorkbookLockedError(
            f"{workbook_path.name} couldn't be updated - close it in Excel and try again. Nothing was changed."
        ) from exc


# --- Transactions ------------------------------------------------------------

def _rng(header: str, last_row: int) -> str:
    return f"{TRANSACTIONS_SHEET}!${COL[header]}$2:${COL[header]}${last_row}"


def _write_header_row(ws: Worksheet, row: int, labels: list[str]) -> None:
    for col, label in enumerate(labels, start=1):
        cell = ws.cell(row=row, column=col, value=label)
        cell.font = BOLD
        cell.fill = HEADER_FILL


def _money(ws: Worksheet, row: int, col: int, value, *, bold: bool = False, total: bool = False):
    cell = ws.cell(row=row, column=col, value=value)
    cell.number_format = CURRENCY_FORMAT
    if bold:
        cell.font = BOLD
    if total:
        cell.border = TOTAL_BORDER
    return cell


def _write_transactions_sheet(wb: Workbook, rows: list[dict]) -> None:
    ws = wb.create_sheet(TRANSACTIONS_SHEET)
    for col, header in enumerate(HEADERS, start=1):
        ws.cell(row=1, column=col, value=header).font = BOLD
        ws.column_dimensions[get_column_letter(col)].width = WIDTHS[col - 1]
    ws.column_dimensions[COL["Row Key"]].hidden = True
    ws.column_dimensions[COL["Suggested Category"]].hidden = True
    ws.freeze_panes = "A2"

    for row_idx, row in enumerate(rows, start=2):
        for col, header in enumerate(HEADERS, start=1):
            cell = ws.cell(row=row_idx, column=col, value=row.get(header))
            if header == "Date" and isinstance(row.get(header), date):
                cell.number_format = "yyyy-mm-dd"
            elif header in ("Debit", "Credit", "Balance", "In/Out Not Shown"):
                cell.number_format = CURRENCY_FORMAT
            elif header == "Confidence":
                cell.number_format = "0.00"
            if row.get("Status") == REVIEW_REQUIRED:
                cell.fill = REVIEW_FILL

    if rows:
        vat_choice = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
        vat_choice.add(f"{COL['VAT']}2:{COL['VAT']}{len(rows) + 1}")
        ws.add_data_validation(vat_choice)

    if not rows:
        return
    last = len(rows) + 1
    totals = last + 2
    ws.cell(row=totals, column=1, value="TOTAL").font = BOLD
    for header in ("Debit", "Credit"):
        col = HEADERS.index(header) + 1
        _money(ws, totals, col, f"=SUM({COL[header]}2:{COL[header]}{last})", bold=True)
    ws.cell(row=totals, column=HEADERS.index("Source File") + 1, value=f"=COUNTA(A2:A{last})").font = Font(
        italic=True
    )

    table = Table(displayName=TABLE_NAME, ref=f"A1:{get_column_letter(len(HEADERS))}{last}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(table)


# --- Shared report inputs ------------------------------------------------------

def report_categories(configured: list[Category], rows: list[dict]) -> list[Category]:
    """Configured categories (plus catch-alls), plus any category typed into
    the sheet by hand that isn't configured - typed as income or expense by
    which way its money mostly flows, so it still lands in the reports.
    """
    categories = with_fallbacks(configured)
    known = {c.name for c in categories}
    extra: dict[str, float] = {}
    for row in rows:
        name = row.get("Category")
        if name and name not in known:
            extra[name] = extra.get(name, 0.0) + (row.get("Credit") or 0) - (row.get("Debit") or 0)
    categories += [Category(name, "income" if net > 0 else "expense") for name, net in sorted(extra.items())]
    return sorted(categories, key=lambda c: TYPE_ORDER.index(c.type))  # stable: config order within type


def _month_balances(rows: list[dict], months: list[str]) -> dict[str, tuple[float | None, float | None]]:
    """(opening, closing) balance per month from the statements' running
    balance. A month's opening is the previous month's closing; the first
    month's is worked back from its first transaction.
    """
    closing: dict[str, float] = {}
    first_opening: float | None = None
    for row in rows:
        balance = row.get("Balance")
        if not row["Month"] or balance is None:
            continue
        if first_opening is None:
            first_opening = balance + (row.get("Debit") or 0) - (row.get("Credit") or 0)
        closing[row["Month"]] = balance

    result = {}
    previous = first_opening
    for month in months:
        result[month] = (previous, closing.get(month))
        previous = closing.get(month, previous)
    return result


# --- Income Statement ------------------------------------------------------------

@dataclass
class IncomeStatementRows:
    total_income: int
    total_expenses: int
    total_transfers: int
    total_drawings: int
    first_month_col: int


def _write_income_statement(
    wb: Workbook, rows: list[dict], categories: list[Category], months: list[str]
) -> IncomeStatementRows:
    ws = wb.create_sheet(INCOME_SHEET)
    last = len(rows) + 1
    debit, credit, cat, month = (_rng(h, last) for h in ("Debit", "Credit", "Category", "Month"))
    first_col = 2
    total_col = first_col + len(months)

    ws["A1"] = "Income Statement (cash basis)"
    ws["A1"].font = TITLE
    ws["A2"] = "From bank transactions by category. Edit config/categories.yaml or a Category cell to reclassify."
    ws["A2"].font = Font(italic=True, color="666666")
    header_row = 4
    # Month headers are the "YYYY-MM" keys the formulas below match on.
    _write_header_row(ws, header_row, ["", *months, "Total"])
    ws.column_dimensions["A"].width = 38
    for col in range(first_col, total_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 14

    row_idx = header_row + 1

    def month_ref(i: int) -> str:
        return f"{get_column_letter(first_col + i)}${header_row}"

    def section(title: str, ctype: str, money_in: bool, total_label: str) -> int:
        nonlocal row_idx
        ws.cell(row=row_idx, column=1, value=title).font = BOLD
        row_idx += 1
        start = row_idx
        for category in (c for c in categories if c.type == ctype):
            ws.cell(row=row_idx, column=1, value=category.name)
            for i in range(len(months)):
                plus, minus = (credit, debit) if money_in else (debit, credit)
                formula = (
                    f"=SUMIFS({plus},{cat},$A{row_idx},{month},{month_ref(i)})"
                    f"-SUMIFS({minus},{cat},$A{row_idx},{month},{month_ref(i)})"
                )
                _money(ws, row_idx, first_col + i, formula)
            _money(ws, row_idx, total_col, _sum_across(row_idx, first_col, total_col - 1), bold=True)
            row_idx += 1
        end = row_idx - 1
        ws.cell(row=row_idx, column=1, value=total_label).font = BOLD
        for col in range(first_col, total_col + 1):
            letter = get_column_letter(col)
            value = f"=SUM({letter}{start}:{letter}{end})" if end >= start else 0
            _money(ws, row_idx, col, value, bold=True, total=True)
        total_row = row_idx
        row_idx += 2
        return total_row

    total_income = section("INCOME", "income", True, "Total income")
    total_expenses = section("EXPENSES", "expense", False, "Total expenses")

    ws.cell(row=row_idx, column=1, value="NET PROFIT / (LOSS)").font = TITLE
    for col in range(first_col, total_col + 1):
        letter = get_column_letter(col)
        cell = _money(ws, row_idx, col, f"={letter}{total_income}-{letter}{total_expenses}", total=True)
        cell.font = Font(bold=True, size=12)
    row_idx += 3

    ws.cell(row=row_idx, column=1, value="Not included in profit (money out, net)").font = Font(bold=True, italic=True)
    row_idx += 1
    total_transfers = section("Transfers to own accounts", "transfer", False, "Total transfers")
    total_drawings = section("Owner drawings / personal", "drawings", False, "Total drawings")

    ws.freeze_panes = ws.cell(row=header_row + 1, column=first_col)
    return IncomeStatementRows(total_income, total_expenses, total_transfers, total_drawings, first_col)


def _sum_across(row: int, first_col: int, last_col: int) -> str | int:
    if last_col < first_col:
        return 0
    return f"=SUM({get_column_letter(first_col)}{row}:{get_column_letter(last_col)}{row})"


# --- Monthly Summary --------------------------------------------------------------

def _write_monthly_summary(
    wb: Workbook, rows: list[dict], months: list[str], balances: dict[str, tuple]
) -> None:
    ws = wb.create_sheet(MONTHLY_SHEET)
    last = len(rows) + 1
    debit, credit, month = (_rng(h, last) for h in ("Debit", "Credit", "Month"))

    ws["A1"] = "Monthly Summary"
    ws["A1"].font = TITLE
    header_row = 3
    labels = ["Month", "Opening Balance", "Money In", "Money Out", "Net", "Closing Balance",
              "Check (should be 0)", "Transactions", "Rows to Review"]
    _write_header_row(ws, header_row, labels)
    for col, width in enumerate([10, 16, 16, 16, 16, 16, 18, 13, 14], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width

    r = header_row
    for r, m in enumerate(months, start=header_row + 1):
        opening, closing = balances[m]
        ws.cell(row=r, column=1, value=m)
        _money(ws, r, 2, opening)
        _money(ws, r, 3, f"=SUMIFS({credit},{month},$A{r})")
        _money(ws, r, 4, f"=SUMIFS({debit},{month},$A{r})")
        _money(ws, r, 5, f"=C{r}-D{r}")
        _money(ws, r, 6, closing)
        # Opening + in - out should land exactly on the statement's closing
        # balance; anything else means a missing or misread transaction.
        _money(ws, r, 7, f'=IF(OR(B{r}="",F{r}=""),"",ROUND(B{r}+C{r}-D{r}-F{r},2))')
        ws.cell(row=r, column=8, value=f"=COUNTIFS({month},$A{r})")
        ws.cell(row=r, column=9, value=f'=COUNTIFS({month},$A{r},{_rng("Status", last)},"{REVIEW_REQUIRED}")')
    last_month_row = r

    total_row = last_month_row + 2
    ws.cell(row=total_row, column=1, value="TOTAL").font = BOLD
    for col in (3, 4, 5):
        letter = get_column_letter(col)
        _money(ws, total_row, col, f"=SUM({letter}{header_row + 1}:{letter}{last_month_row})", bold=True, total=True)
    ws.cell(row=total_row, column=8, value=f"=SUM(H{header_row + 1}:H{last_month_row})").font = BOLD
    ws.cell(row=total_row, column=9, value=f"=SUM(I{header_row + 1}:I{last_month_row})").font = BOLD
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)

    if months:
        chart = BarChart()
        chart.title = "Money in vs money out"
        chart.y_axis.title = "Rand"
        chart.y_axis.numFmt = '"R" #,##0'
        data = Reference(ws, min_col=3, max_col=4, min_row=header_row, max_row=last_month_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=header_row + 1, max_row=last_month_row))
        chart.height, chart.width = 8, 18
        ws.add_chart(chart, f"A{total_row + 3}")


# --- Cash Flow --------------------------------------------------------------------

def _write_cash_flow(
    wb: Workbook, months: list[str], balances: dict[str, tuple], income: IncomeStatementRows
) -> None:
    ws = wb.create_sheet(CASHFLOW_SHEET)
    first_col = 2
    total_col = first_col + len(months)
    src = f"'{INCOME_SHEET}'!"

    ws["A1"] = "Cash Flow Statement"
    ws["A1"].font = TITLE
    header_row = 3
    _write_header_row(ws, header_row, ["", *months, "Total"])
    ws.column_dimensions["A"].width = 40
    for col in range(first_col, total_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 14

    lines = [
        ("opening", "Opening balance", True),
        ("h_ops", "Operating activities", True),
        ("income", "  Income received", False),
        ("expenses", "  Expenses paid", False),
        ("ops", "Net cash from operations", True),
        ("h_other", "Other movements", True),
        ("transfers", "  Transfers to own accounts", False),
        ("drawings", "  Owner drawings / personal", False),
        ("change", "Net change in cash", True),
        ("closing_calc", "Closing balance (calculated)", True),
        ("closing_stmt", "Closing balance per bank statement", False),
        ("check", "Difference (should be 0)", False),
    ]
    row_of = {}
    for offset, (key, label, bold) in enumerate(lines, start=1):
        row_of[key] = header_row + offset
        ws.cell(row=row_of[key], column=1, value=label).font = Font(bold=bold)

    def put(key: str, col: int, value, **kw):
        if not key.startswith("h_"):
            _money(ws, row_of[key], col, value, **kw)

    for i, m in enumerate(months):
        col = first_col + i
        letter = get_column_letter(col)
        is_letter = get_column_letter(income.first_month_col + i)
        opening, closing = balances[m]
        _cash_flow_column(put, row_of, col, letter, opening, closing, src, is_letter, income)

    # Total column: flows sum across months; balances are the first opening
    # and last closing.
    col, letter = total_col, get_column_letter(total_col)
    first_opening = balances[months[0]][0] if months else None
    last_closing = balances[months[-1]][1] if months else None
    is_total_letter = get_column_letter(income.first_month_col + len(months))
    _cash_flow_column(put, row_of, col, letter, first_opening, last_closing, src, is_total_letter, income)
    ws.freeze_panes = ws.cell(row=header_row + 1, column=first_col)


def _cash_flow_column(put, row_of, col, letter, opening, closing, src, is_letter, income: IncomeStatementRows):
    put("opening", col, opening, bold=True)
    put("income", col, f"={src}{is_letter}{income.total_income}")
    put("expenses", col, f"=-{src}{is_letter}{income.total_expenses}")
    put("ops", col, f"={letter}{row_of['income']}+{letter}{row_of['expenses']}", bold=True, total=True)
    put("transfers", col, f"=-{src}{is_letter}{income.total_transfers}")
    put("drawings", col, f"=-{src}{is_letter}{income.total_drawings}")
    put("change", col, f"={letter}{row_of['ops']}+{letter}{row_of['transfers']}+{letter}{row_of['drawings']}",
        bold=True, total=True)
    put("closing_calc", col, f"={letter}{row_of['opening']}+{letter}{row_of['change']}", bold=True)
    put("closing_stmt", col, closing)
    put("check", col, f'=IF({letter}{row_of["closing_stmt"]}="","",'
                      f'ROUND({letter}{row_of["closing_calc"]}-{letter}{row_of["closing_stmt"]},2))')


# --- Category Breakdown --------------------------------------------------------------

def _write_category_breakdown(wb: Workbook, rows: list[dict], categories: list[Category]) -> None:
    ws = wb.create_sheet(BREAKDOWN_SHEET)
    last = len(rows) + 1
    debit, credit, cat = (_rng(h, last) for h in ("Debit", "Credit", "Category"))

    ws["A1"] = "Category Breakdown"
    ws["A1"].font = TITLE
    header_row = 3
    _write_header_row(ws, header_row, ["Category", "Type", "Transactions", "Money In", "Money Out", "Net",
                                       "% of Expenses"])
    for col, width in enumerate([38, 10, 13, 15, 15, 15, 14], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width

    first = header_row + 1
    last_cat = first + len(categories) - 1
    for r, category in enumerate(categories, start=first):
        ws.cell(row=r, column=1, value=category.name)
        ws.cell(row=r, column=2, value=category.type)
        ws.cell(row=r, column=3, value=f"=COUNTIFS({cat},$A{r})")
        _money(ws, r, 4, f"=SUMIFS({credit},{cat},$A{r})")
        _money(ws, r, 5, f"=SUMIFS({debit},{cat},$A{r})")
        _money(ws, r, 6, f"=D{r}-E{r}")
        pct = ws.cell(
            row=r, column=7,
            value=f'=IF(AND(B{r}="expense",$E${last_cat + 2}<>0),(E{r}-D{r})/$E${last_cat + 2},"")',
        )
        pct.number_format = PERCENT_FORMAT

    total_row = last_cat + 2
    ws.cell(row=total_row, column=1, value="Total expenses").font = BOLD
    _money(ws, total_row, 5,
           f'=SUMIFS(E{first}:E{last_cat},B{first}:B{last_cat},"expense")'
           f'-SUMIFS(D{first}:D{last_cat},B{first}:B{last_cat},"expense")', bold=True, total=True)

    expense_rows = [first + i for i, c in enumerate(categories) if c.type == "expense"]
    if expense_rows:
        chart = BarChart()
        chart.type = "bar"
        chart.title = "Spending by category"
        chart.legend = None
        chart.x_axis.numFmt = '"R" #,##0'
        chart.add_data(Reference(ws, min_col=5, min_row=expense_rows[0], max_row=expense_rows[-1]))
        chart.set_categories(Reference(ws, min_col=1, min_row=expense_rows[0], max_row=expense_rows[-1]))
        chart.height, chart.width = 9, 18
        ws.add_chart(chart, f"I{header_row}")


# --- Monthly income & expenses with VAT --------------------------------------------------

VAT_SHEET_HEADERS = ["Date", "Reference", "Description", "Category", "VAT (Yes/No)", "Amount",
                     "Amount excl VAT", f"VAT {VAT_RATE:.0%} (calculated)"]
VAT_SHEET_WIDTHS = [11, 11, 48, 30, 12, 14, 16, 13]
AMOUNT_COL, EXCL_COL, VAT_COL = 6, 7, 8  # F, G, H on each month sheet
# Hidden helper columns: the transaction's Row Key, and the row it's on in
# the Transactions sheet right now (looked up, so sorting or filtering
# Transactions in Excel can't point a month sheet at the wrong transaction;
# a deleted transaction shows #N/A instead of a wrong number).
KEY_COL, MATCH_COL = 9, 10  # I, J


@dataclass
class MonthVatTotals:
    sheet: str
    income_total_row: int
    expense_total_row: int


def _write_month_vat_sheets(
    wb: Workbook, rows: list[dict], categories: list[Category], months: list[str]
) -> list[MonthVatTotals]:
    """One sheet per month: INCOME and EXPENSES listed separately, each
    amount split into excl-VAT and VAT by the transaction's VAT Yes/No.
    Cells look the transaction up in the Transactions sheet by Row Key, so
    edits there flow through.
    """
    types = {c.name: c.type for c in categories}
    clients = [r["Client"] for r in rows if r.get("Client")]
    client = max(set(clients), key=clients.count) if clients else ""
    divisor = f"{1 + VAT_RATE:g}"

    key_range = f"{TRANSACTIONS_SHEET}!${COL['Row Key']}:${COL['Row Key']}"

    def tx(header: str, r: int) -> str:
        return f"INDEX({TRANSACTIONS_SHEET}!${COL[header]}:${COL[header]},$J{r})"

    totals = []
    for month in months:
        start = datetime.strptime(month, "%Y-%m").date()
        end = date(start.year, start.month, calendar.monthrange(start.year, start.month)[1])
        name = start.strftime("%b %Y")
        ref_prefix = start.strftime("%b%y").upper()
        ws = wb.create_sheet(name)
        for col, width in enumerate(VAT_SHEET_WIDTHS, start=1):
            ws.column_dimensions[get_column_letter(col)].width = width
        for col in (KEY_COL, MATCH_COL):
            ws.column_dimensions[get_column_letter(col)].hidden = True

        ws["A1"] = f"{client} | Bank statement income and expenses" if client else "Bank statement income and expenses"
        ws["A2"] = (f"Income & Expense Report for {start:%B %Y} "
                    f"({start:%d/%m/%y} - {end:%d/%m/%y})")
        ws["A2"].font = TITLE

        month_rows = [row for row in rows if row["Month"] == month]
        r = 4

        def section(title: str, ctype: str, ref_letter: str, total_label: str) -> int:
            nonlocal r
            ws.cell(row=r, column=1, value=title).font = BOLD
            r += 1
            _write_header_row(ws, r, VAT_SHEET_HEADERS)
            r += 1
            first = r
            picked = [row for row in month_rows if types.get(row["Category"]) == ctype]
            for n, row in enumerate(picked, start=1):
                ws.cell(row=r, column=KEY_COL, value=row["Row Key"])
                ws.cell(row=r, column=MATCH_COL, value=f"=MATCH($I{r},{key_range},0)")
                ws.cell(row=r, column=1, value=row["Date"]).number_format = "dd/mm/yy"
                ws.cell(row=r, column=2, value=f"{ref_prefix}{ref_letter}{n:02d}")
                ws.cell(row=r, column=3, value=row.get("Description"))
                ws.cell(row=r, column=4, value=f"={tx('Category', r)}")
                ws.cell(row=r, column=5, value=f"={tx('VAT', r)}")
                money_in = f"{tx('Credit', r)}-{tx('Debit', r)}"
                money_out = f"{tx('Debit', r)}-{tx('Credit', r)}"
                _money(ws, r, AMOUNT_COL, f"={money_in if ctype == 'income' else money_out}")
                _money(ws, r, EXCL_COL, f'=IF(E{r}="Yes",F{r}/{divisor},F{r})')
                _money(ws, r, VAT_COL, f"=F{r}-G{r}")
                r += 1
            last = r - 1
            ws.cell(row=r, column=1, value=total_label).font = BOLD
            for col in (AMOUNT_COL, EXCL_COL, VAT_COL):
                letter = get_column_letter(col)
                value = f"=SUM({letter}{first}:{letter}{last})" if last >= first else 0
                cell = _money(ws, r, col, value, bold=True, total=True)
                cell.fill = TOTAL_FILL
            total_row = r
            r += 2
            return total_row

        income_total = section("INCOME", "income", "R", "Total income")
        expense_total = section("EXPENSES", "expense", "P", "Total expenses")

        r += 1
        ws.cell(row=r, column=1, value=f"{start:%B %Y} summary").font = BOLD
        r += 1
        for col, label in ((AMOUNT_COL, "Total"), (EXCL_COL, "Excl VAT"), (VAT_COL, "VAT")):
            cell = ws.cell(row=r, column=col, value=label)
            cell.font = BOLD
            cell.fill = TOTAL_FILL
        r += 1
        summary_lines = [
            ("Total income", lambda L: f"={L}{income_total}"),
            ("Total expenses", lambda L: f"={L}{expense_total}"),
            ("Difference", lambda L: f"={L}{income_total}-{L}{expense_total}"),
        ]
        for label, formula in summary_lines:
            ws.cell(row=r, column=1, value=label).font = BOLD if label == "Difference" else Font()
            for col in (AMOUNT_COL, EXCL_COL, VAT_COL):
                _money(ws, r, col, formula(get_column_letter(col)), bold=label == "Difference",
                       total=label == "Difference")
            r += 1

        totals.append(MonthVatTotals(name, income_total, expense_total))
    return totals


# The SARS VAT201 fields the bank statements can support, each worked out
# from the VAT Summary sheet's columns (B income, D VAT on income, G VAT on
# expenses) of the month's row. Amounts with VAT include it, as on the
# return; VAT is the 15/115 part of them.
VAT201_SHEET = "VAT201"
VAT201_FIELDS = [
    ("1", "Standard-rated supplies (income with VAT), incl VAT", "=VS!D{r}*(1+{rate})/{rate}"),
    ("2 / 3", "Income without VAT - zero-rated (field 2) or exempt (field 3): put it in the right one",
     "=VS!B{r}-VS!D{r}*(1+{rate})/{rate}"),
    ("4", "Output tax on field 1", "=VS!D{r}"),
    ("13", "Total output tax", "=VS!D{r}"),
    ("15", "Input tax on goods and services (capital goods belong in field 14)", "=VS!G{r}"),
    ("19", "Total input tax", "=VS!G{r}"),
    ("20", "VAT payable (negative: refundable)", "=VS!D{r}-VS!G{r}"),
]


def vat201_by_month(months: list[dict]) -> list[dict]:
    """The VAT201 fields per month from vat_by_month's figures, as on the
    VAT201 sheet (not rounded)."""
    return [{"Month": m["Month"], "1": m["Income with VAT"], "2 / 3": m["Income"] - m["Income with VAT"],
             "4": m["VAT on income"], "13": m["VAT on income"], "15": m["VAT on expenses"],
             "19": m["VAT on expenses"], "20": m["VAT on income"] - m["VAT on expenses"]} for m in months]


def _write_vat201(wb: Workbook, months: list[MonthVatTotals]) -> None:
    ws = wb.create_sheet(VAT201_SHEET)
    ws["A1"] = "VAT201 - figures for the SARS return, per month"
    ws["A1"].font = TITLE
    ws["A2"] = (f"CALCULATED at {VAT_RATE:.0%} from each transaction's VAT Yes/No - not read from the statements. "
                "Add up the months of your VAT period (e.g. two months for category A or B). Only fields bank "
                "statements can support are here; claim input tax only where you hold a valid tax invoice.")
    ws["A2"].font = Font(italic=True, color="666666")
    header_row = 4
    _write_header_row(ws, header_row, ["Field", "Description", *(m.sheet for m in months), "Total"])
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 70
    for col in range(3, len(months) + 4):
        ws.column_dimensions[get_column_letter(col)].width = 14
    summary = f"'{VAT_SUMMARY_SHEET}'"
    for r, (field, label, formula) in enumerate(VAT201_FIELDS, start=header_row + 1):
        ws.cell(row=r, column=1, value=field).font = BOLD
        ws.cell(row=r, column=2, value=label)
        for i in range(len(months)):
            # The VAT Summary lists the months from its row 5, in this order.
            cell_formula = formula.format(r=5 + i, rate=f"{VAT_RATE:g}").replace("VS!", f"{summary}!")
            _money(ws, r, 3 + i, cell_formula, bold=field == "20")
        first, last = get_column_letter(3), get_column_letter(2 + len(months))
        cell = _money(ws, r, 3 + len(months), f"=SUM({first}{r}:{last}{r})", bold=True)
        cell.fill = TOTAL_FILL
    ws.freeze_panes = ws.cell(row=header_row + 1, column=3)


def _write_vat_summary(wb: Workbook, months: list[MonthVatTotals]) -> None:
    ws = wb.create_sheet(VAT_SUMMARY_SHEET)
    ws["A1"] = "Income, Expenses and VAT per Month"
    ws["A1"].font = TITLE
    ws["A2"] = (f"VAT is CALCULATED at {VAT_RATE:.0%} as instructed - it is not read from the statements. "
                f"VAT payable = VAT on income minus VAT on expenses; "
                "negative means a refund. Set each transaction's VAT Yes/No on the Transactions sheet.")
    ws["A2"].font = Font(italic=True, color="666666")
    header_row = 4
    _write_header_row(ws, header_row, ["Month", "Income", "Income excl VAT", "VAT on income", "Expenses",
                                       "Expenses excl VAT", "VAT on expenses", "Difference",
                                       "Difference excl VAT", "VAT payable"])
    for col, width in enumerate([12, 15, 16, 15, 15, 17, 16, 15, 19, 15], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width

    r = header_row
    amount, excl, vat = (get_column_letter(c) for c in (AMOUNT_COL, EXCL_COL, VAT_COL))
    for r, m in enumerate(months, start=header_row + 1):
        src = f"'{m.sheet}'!"
        ws.cell(row=r, column=1, value=m.sheet)
        for col, (letter, total_row) in enumerate(
            [(amount, m.income_total_row), (excl, m.income_total_row), (vat, m.income_total_row),
             (amount, m.expense_total_row), (excl, m.expense_total_row), (vat, m.expense_total_row)],
            start=2,
        ):
            _money(ws, r, col, f"={src}{letter}{total_row}")
        _money(ws, r, 8, f"=B{r}-E{r}")
        _money(ws, r, 9, f"=C{r}-F{r}")
        _money(ws, r, 10, f"=D{r}-G{r}", bold=True)
    last = r

    total_row = last + 2
    ws.cell(row=total_row, column=1, value="TOTAL").font = BOLD
    for col in range(2, 11):
        letter = get_column_letter(col)
        cell = _money(ws, total_row, col, f"=SUM({letter}{header_row + 1}:{letter}{last})", bold=True, total=True)
        cell.fill = TOTAL_FILL
    ws.freeze_panes = ws.cell(row=header_row + 1, column=2)


# --- Categories reference ----------------------------------------------------------------

def _write_categories_sheet(wb: Workbook, categories: list[Category]) -> None:
    ws = wb.create_sheet(CATEGORIES_SHEET)
    ws["A1"] = "Categories (from config/categories.yaml - edit that file to change the rules)"
    ws["A1"].font = BOLD
    _write_header_row(ws, 3, ["Category", "Type", "Matches descriptions containing"])
    for col, width in enumerate([38, 10, 80], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    for r, category in enumerate(categories, start=4):
        ws.cell(row=r, column=1, value=category.name)
        ws.cell(row=r, column=2, value=category.type)
        ws.cell(row=r, column=3, value=", ".join(category.match) or "(anything not matched above)")


# --- Review (one line per statement) ------------------------------------------------------

REVIEW_COLUMNS = [
    "Source File", "Statement Period", "Status", "Opening Balance", "Total Credits", "Total Debits",
    "Net Movement", "Closing (computed)", "Closing (printed)", "Printed Credits", "Printed Debits",
    "Printed Fees", "Rows Needing Review", "Issues",
]
REVIEW_WIDTHS = [34, 26, 17, 15, 15, 15, 15, 17, 16, 15, 15, 13, 12, 90]
REVIEW_HEADER_ROW = 5


def statement_review_row(result) -> dict:
    """A statement's line on the Review sheet, from its StatementResult."""
    report = result.report or {}
    return {
        "Source File": result.source_file,
        "Statement Period": result.statement_period,
        "Status": result.status,
        "Opening Balance": report.get("opening_balance"),
        "Total Credits": report.get("total_credits"),
        "Total Debits": report.get("total_debits"),
        "Net Movement": report.get("net_movement"),
        "Closing (computed)": report.get("computed_closing"),
        "Closing (printed)": report.get("printed_closing"),
        "Printed Credits": report.get("printed_total_credits"),
        "Printed Debits": report.get("printed_total_debits"),
        "Printed Fees": report.get("printed_total_fees"),
        "Rows Needing Review": sum(1 for t in result.transactions if t.status == REVIEW_REQUIRED),
        "Issues": "; ".join(result.problems) or "",
    }


def read_reviews(workbook_path: Path) -> list[dict]:
    """The Review sheet's line per statement, as {header: value}."""
    wb = _open_workbook(workbook_path)
    try:
        return _read_reviews(wb)
    finally:
        wb.close()


def vat_by_month(rows: list[dict], categories: list[Category]) -> list[dict]:
    """Per month, the figures the VAT Summary sheet calculates: income and
    expenses, and the VAT in those whose VAT cell is Yes (calculated at
    VAT_RATE, never read from the statements). Not rounded - like the
    sheet's formulas - so totals over several months come out the same as
    the workbook's; round only to show them."""
    types = {c.name: c.type for c in report_categories(categories, rows)}
    months: dict[str, dict] = {}
    for row in rows:
        kind = types.get(row.get("Category"))
        if kind not in ("income", "expense") or not row.get("Month"):
            continue
        amount = (row.get("Credit") or 0) - (row.get("Debit") or 0)
        if kind == "expense":
            amount = -amount
        with_vat = row.get("VAT") == "Yes"
        m = months.setdefault(row["Month"], {"Month": row["Month"], "Income": 0.0, "VAT on income": 0.0,
                                             "Expenses": 0.0, "VAT on expenses": 0.0, "Income with VAT": 0.0})
        label = "Income" if kind == "income" else "Expenses"
        m[label] += amount
        m[f"VAT on {label.lower()}"] += amount - amount / (1 + VAT_RATE) if with_vat else 0.0
        if with_vat and kind == "income":
            m["Income with VAT"] += amount
    out = []
    for month in sorted(months):
        m = months[month]
        m["VAT payable"] = m["VAT on income"] - m["VAT on expenses"]
        out.append(m)
    return out


def _read_reviews(wb: Workbook) -> list[dict]:
    if REVIEW_SHEET not in wb.sheetnames:
        return []
    ws = wb[REVIEW_SHEET]
    rows = list(ws.iter_rows(values_only=True))
    header_at = next((i for i, r in enumerate(rows) if r and r[0] == "Source File"), None)
    if header_at is None:
        return []
    headers = rows[header_at]
    return [{h: v for h, v in zip(headers, r) if h} for r in rows[header_at + 1:] if r and r[0]]


def _write_review_sheet(wb: Workbook, reviews: list[dict], transaction_count: int) -> None:
    ws = wb.create_sheet(REVIEW_SHEET)
    ws["A1"] = "Statement Review"
    ws["A1"].font = TITLE
    ws["A2"] = ("APPROVED: every row is confirmed by the statement's own running balance and by a second, "
                "independent reading, and opening + credits - debits = closing. REVIEW_REQUIRED: something "
                "couldn't be confirmed - see Issues, and the highlighted rows on the Transactions sheet.")
    ws["A2"].font = Font(italic=True, color="666666")
    ws["A3"] = "Transactions needing review:"
    ws["A3"].font = BOLD
    last = transaction_count + 1
    ws["C3"] = f'=COUNTIF({_rng("Status", last)},"{REVIEW_REQUIRED}")' if transaction_count else 0
    ws["C3"].font = BOLD
    _write_header_row(ws, REVIEW_HEADER_ROW, REVIEW_COLUMNS)
    for col, width in enumerate(REVIEW_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    for r, review in enumerate(reviews, start=REVIEW_HEADER_ROW + 1):
        for c, header in enumerate(REVIEW_COLUMNS, start=1):
            cell = ws.cell(row=r, column=c, value=review.get(header))
            if header in ("Opening Balance", "Total Credits", "Total Debits", "Net Movement", "Closing (computed)",
                          "Closing (printed)", "Printed Credits", "Printed Debits", "Printed Fees"):
                cell.number_format = CURRENCY_FORMAT
        status_cell = ws.cell(row=r, column=REVIEW_COLUMNS.index("Status") + 1)
        status_cell.fill = REVIEW_FILL if review.get("Status") == REVIEW_REQUIRED else APPROVED_FILL
        status_cell.font = BOLD
    ws.freeze_panes = ws.cell(row=REVIEW_HEADER_ROW + 1, column=2)
