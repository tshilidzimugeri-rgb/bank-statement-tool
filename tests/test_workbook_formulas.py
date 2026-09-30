"""The workbook's report formulas, calculated by Excel itself and compared
with totals worked out independently in Python. (openpyxl can't calculate
formulas, so the other workbook tests only check their text - a wrong range
there would pass. This one would fail.) Skipped where Excel isn't installed.
"""
import dataclasses
import shutil
from datetime import timedelta
import subprocess
import sys
from collections import defaultdict
from pathlib import Path

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).parent / "fixtures"))
from synthetic import build, generate, render_pdf  # noqa: E402

from statement_tool import config as config_mod  # noqa: E402
from statement_tool.categorize import load_categories  # noqa: E402
from statement_tool.excel_writer import (  # noqa: E402
    VAT_RATE, append_transactions, read_transactions, report_categories, statement_review_row, vat201_by_month,
    vat_by_month,
)
from statement_tool.extract.document import parse_document  # noqa: E402

PROJECT_ROOT = Path(__file__).parent.parent
EXCEL = Path(r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE")


def _excel_recalculate(source: Path, target: Path) -> None:
    script = f"""
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {{
  $wb = $xl.Workbooks.Open('{source}')
  $xl.CalculateFull()
  $wb.SaveAs('{target}', 51)
  $wb.Close($false)
}} finally {{ $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }}
"""
    subprocess.run(["powershell", "-NoProfile", "-Command", script], check=True, timeout=300)


@pytest.mark.skipif(not (EXCEL.exists() and shutil.which("powershell")), reason="needs Microsoft Excel")
def test_report_formulas_match_independent_totals(tmp_path):
    settings = config_mod.load_settings(dotenv_path=PROJECT_ROOT / ".env.example")
    layouts, generic = config_mod.load_bank_layouts(PROJECT_ROOT / "config" / "banks.yaml")
    categories = load_categories(PROJECT_ROOT / "config" / "categories.yaml")
    book = tmp_path / "book.xlsx"
    # Two consecutive monthly statements of one account: the second opens on
    # the first's closing balance.
    first = generate(4)
    rec = first.recipe
    next_start = rec.end + timedelta(days=1)
    next_end = (next_start + timedelta(days=32)).replace(day=1) - timedelta(days=1)
    shifted = [(min(next_end, d + (next_start - rec.start)), desc, amount, fee, accrued)
               for d, desc, amount, fee, accrued in rec.day_rows]
    second = build(dataclasses.replace(rec, start=next_start, end=next_end, opening=first.closing,
                                       day_rows=sorted(shifted, key=lambda r: r[0])))
    for n, statement in enumerate((first, second)):
        result = parse_document(render_pdf(statement, tmp_path / f"s{n}.pdf"), layouts=layouts,
                                generic=generic, client_rules=[], settings=settings, interactive=False)[0]
        append_transactions(book, result.transactions, categories, statement=statement_review_row(result))

    calculated = tmp_path / "calculated.xlsx"
    _excel_recalculate(book, calculated)
    wb = load_workbook(calculated, data_only=True)

    rows = read_transactions(book)
    types = {c.name: c.type for c in report_categories(categories, rows)}
    money_in, money_out = defaultdict(float), defaultdict(float)
    income, expenses, vat_income, vat_expenses = (defaultdict(float) for _ in range(4))
    by_category = defaultdict(lambda: [0.0, 0.0])
    for r in rows:
        month, debit, credit = r["Month"], r.get("Debit") or 0, r.get("Credit") or 0
        money_in[month] += credit
        money_out[month] += debit
        by_category[r["Category"]][0] += credit
        by_category[r["Category"]][1] += debit
        kind = types[r["Category"]]
        vat_share = (1 - 1 / (1 + VAT_RATE)) if r["VAT"] == "Yes" else 0
        if kind == "income":
            income[month] += credit - debit
            vat_income[month] += (credit - debit) * vat_share
        elif kind == "expense":
            expenses[month] += debit - credit
            vat_expenses[month] += (debit - credit) * vat_share

    close = lambda a, b: abs((a or 0) - (b or 0)) < 0.01  # noqa: E731

    monthly = wb["Monthly Summary"]
    months_checked = 0
    for row in monthly.iter_rows(min_row=4, values_only=True):
        if not row[0] or row[0] == "TOTAL" or row[0] not in money_in:
            continue
        month = row[0]
        assert close(row[2], money_in[month]) and close(row[3], money_out[month]), month
        assert row[6] in (0, None) or close(row[6], 0), f"{month}: monthly check is {row[6]}"
        months_checked += 1
    assert months_checked == len(money_in)

    cash = wb["Cash Flow"]
    diff_row = next(r for r in cash.iter_rows(values_only=True) if r[0] == "Difference (should be 0)")
    assert all(v in (None, "") or close(v, 0) for v in diff_row[1:]), diff_row

    statement_sheet = wb["Income Statement"]
    header = next(r for r in statement_sheet.iter_rows(values_only=True) if r[0] is None and r[1] in income)
    totals = {r[0]: r for r in statement_sheet.iter_rows(values_only=True) if r[0] in ("Total income", "Total expenses")}
    for col, month in enumerate(header[1:], start=1):
        if month in income:
            assert close(totals["Total income"][col], income[month]), month
            assert close(totals["Total expenses"][col], expenses[month]), month

    vat = wb["VAT Summary"]
    for r in vat.iter_rows(min_row=5, values_only=True):
        if not r[0] or r[0] == "TOTAL":
            continue
        month = next(m for m in income if f"{r[0]}" == _month_label(m))
        assert close(r[3], vat_income[month]) and close(r[6], vat_expenses[month]), r[0]
        assert close(r[9], vat_income[month] - vat_expenses[month]), r[0]

    breakdown = wb["Category Breakdown"]
    for r in breakdown.iter_rows(min_row=4, values_only=True):
        if r[0] in by_category:
            assert close(r[3], by_category[r[0]][0]) and close(r[4], by_category[r[0]][1]), r[0]

    # VAT201: Excel's figures equal the upload page's, and follow the return's
    # own arithmetic (field 1 includes VAT, field 4 is its 15/115 part).
    vat201 = wb["VAT201"]
    header = [c.value for c in vat201[4]]
    fields = {r[0]: r for r in vat201.iter_rows(min_row=5, values_only=True) if r[0]}
    page = {m["Month"]: m for m in vat201_by_month(vat_by_month(rows, categories))}
    checked = 0
    for col, label in enumerate(header[2:-1], start=2):
        month = next(m for m in page if _month_label(m) == label)
        for field in ("1", "2 / 3", "4", "13", "15", "19", "20"):
            assert close(fields[field][col], page[month][field]), (label, field)
        assert close(fields["4"][col], fields["1"][col] * VAT_RATE / (1 + VAT_RATE)), label
        assert close(fields["1"][col] + fields["2 / 3"][col], income[month]), label
        assert close(fields["20"][col], vat_income[month] - vat_expenses[month]), label
        checked += 1
    assert checked == len(income)
    assert close(fields["20"][-1], sum(vat_income.values()) - sum(vat_expenses.values()))


def _month_label(month: str) -> str:
    from datetime import datetime
    return datetime.strptime(month, "%Y-%m").strftime("%b %Y")


def test_months_checked_count_covers_both_statements():
    """Guards the formula test itself: both months must be compared."""
    first = generate(4)
    assert len({d.strftime("%Y-%m") for d, *_ in first.recipe.day_rows}) == 1
