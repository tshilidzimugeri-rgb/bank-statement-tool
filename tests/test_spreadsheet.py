"""Bank statements as messy spreadsheets (made up from fixtures/synthetic.py):
titles and notes around the table, headings in different words, amounts as
text, debit/credit columns or a signed amount or a Cr/Dr column, CSV with
semicolons and decimal commas. Each must be read exactly as in the file,
and a changed figure must never be APPROVED."""
import csv
import sys
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).parent / "fixtures"))
from synthetic import generate  # noqa: E402

from statement_tool import config as config_mod  # noqa: E402
from statement_tool.extract.spreadsheet import parse_spreadsheet  # noqa: E402
from statement_tool.periods import periods, vat_periods  # noqa: E402

PROJECT_ROOT = Path(__file__).parent.parent
SETTINGS = config_mod.load_settings(dotenv_path=PROJECT_ROOT / ".env.example")
LAYOUTS, GENERIC = config_mod.load_bank_layouts(PROJECT_ROOT / "config" / "banks.yaml")


def _truth(seed):
    st = generate(seed)
    return st, [t for t in st.truth]


def _text_money(value: float, comma: bool = False) -> str:
    text = f"{abs(value):,.2f}"
    if comma:
        text = text.replace(",", " ").replace(".", ",")
    return text


def _read(path):
    results = parse_spreadsheet(path, layouts=LAYOUTS, generic=GENERIC, client_rules=[], settings=SETTINGS)
    return results, [t for r in results for t in r.transactions]


def _figures(rows):
    return [(t.date, t.debit, t.credit, t.balance) for t in rows]


def _xlsx(path, title_rows, header, body, footer=()):
    wb = Workbook()
    ws = wb.active
    for r in title_rows:
        ws.append(r)
    ws.append(header)
    for r in body:
        ws.append(r)
    for r in footer:
        ws.append(r)
    wb.save(path)
    return path


@pytest.mark.parametrize("seed", range(4))
def test_debit_credit_columns_with_titles_totals_and_extra_columns(seed, tmp_path):
    st, truth = _truth(seed)
    body = [[date.fromisoformat(t.date), f"Item {i}", f"REF{i}", t.debit, t.credit, t.balance, "p-1"]
            for i, t in enumerate(truth)]
    path = _xlsx(tmp_path / "s.xlsx",
                 [["Harbour Mutual Bank - Account statement"], [], [f"Account number {st.account}"], []],
                 ["Transaction Date", "Details", "Reference", "Money Out", "Money In", "Running Balance", "Page"],
                 [[date.fromisoformat(truth[0].date), "Opening balance", None, None, None, st.opening, None], *body],
                 [[], [None, "Totals", None, sum(t.debit or 0 for t in truth), sum(t.credit or 0 for t in truth)]])
    results, rows = _read(path)
    assert _figures(rows) == [(t.date, t.debit, t.credit, t.balance) for t in truth]
    assert results[0].account_number == st.account
    assert {t.status for t in rows} == {"APPROVED"}


@pytest.mark.parametrize("seed", range(4))
def test_signed_amounts_typed_as_text_and_dates_as_text(seed, tmp_path):
    st, truth = _truth(seed)
    body = []
    for i, t in enumerate(truth):
        amount = f"R-{_text_money(t.debit)}" if t.debit is not None else f"R{_text_money(t.credit)}"
        balance = None if t.balance is None else (f"-R{_text_money(t.balance)}" if t.balance < 0
                                                  else f"R{_text_money(t.balance)}")
        d = date.fromisoformat(t.date)
        body.append([f"{d:%d/%m/%Y}", f"Payment {i}", amount, balance])
    path = _xlsx(tmp_path / "s.xlsx", [[f"Opening Balance: R{_text_money(st.opening)}"], []],
                 ["Date", "Description", "Amount", "Balance"], body)
    results, rows = _read(path)
    assert _figures(rows) == [(t.date, t.debit, t.credit, t.balance) for t in truth]
    assert {t.status for t in rows} == {"APPROVED"}


@pytest.mark.parametrize("seed", range(4))
def test_csv_with_semicolons_decimal_commas_and_a_cr_dr_column(seed, tmp_path):
    st, truth = _truth(seed)
    path = tmp_path / "s.csv"
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Statement export"])
        w.writerow([])
        w.writerow(["Opening balance", "", "", "", _text_money(st.opening, comma=True)])
        w.writerow(["Date", "Narrative", "Amount", "Dr/Cr", "Balance"])
        for i, t in enumerate(truth):
            amount = t.debit if t.debit is not None else t.credit
            balance = "" if t.balance is None else ("-" if t.balance < 0 else "") + _text_money(t.balance, comma=True)
            w.writerow([t.date, f"Item; with a semicolon {i}", _text_money(amount, comma=True),
                        "Dr" if t.debit is not None else "Cr", balance])
    results, rows = _read(path)
    assert _figures(rows) == [(t.date, t.debit, t.credit, t.balance) for t in truth]
    assert {t.status for t in rows} == {"APPROVED"}


@pytest.mark.parametrize("seed", range(4))
def test_a_changed_amount_is_never_approved(seed, tmp_path):
    st, truth = _truth(seed)
    changed = next(i for i, t in enumerate(truth) if t.balance is not None and i > 0)
    body = []
    for i, t in enumerate(truth):
        debit = t.debit if i != changed or t.debit is None else round(t.debit + 10, 2)
        credit = t.credit if i != changed or t.credit is None else round(t.credit + 10, 2)
        body.append([date.fromisoformat(t.date), f"Item {i}", debit, credit, t.balance])
    path = _xlsx(tmp_path / "s.xlsx", [], ["Date", "Description", "Debit", "Credit", "Balance"],
                 [[date.fromisoformat(truth[0].date), "Balance brought forward", None, None, st.opening], *body])
    results, rows = _read(path)
    truth_set = {(t.date, t.debit, t.credit, t.balance) for t in truth}
    assert [t for t in rows if t.status == "APPROVED" and (t.date, t.debit, t.credit, t.balance) not in truth_set] == []
    assert any(t.status == "REVIEW_REQUIRED" for t in rows)
    assert results[0].status == "REVIEW_REQUIRED"


def test_a_file_without_a_transaction_table_is_explained(tmp_path):
    path = _xlsx(tmp_path / "s.xlsx", [["Just some notes"]], ["Name", "Phone"], [["Jane", "012 345 6789"]])
    results, rows = _read(path)
    assert rows == [] and not results[0].ok and "no table of transactions" in results[0].error


def test_reporting_periods_follow_the_march_to_february_financial_year():
    labels = [p.label for p in periods(["2026-01", "2026-02", "2026-03", "2026-04"])]
    assert labels == [
        "All months (Jan 2026 - Apr 2026)",
        "FY2025/26 (Mar 2025 - Feb 2026)", "FY2025/26 Q4 (Dec - Feb 2026)",
        "FY2026/27 (Mar 2026 - Feb 2027)", "FY2026/27 Q1 (Mar - May 2026)",
        "Jan 2026", "Feb 2026", "Mar 2026", "Apr 2026",
    ]
    q4 = next(p for p in periods(["2026-01", "2026-02"]) if "Q4" in p.label)
    assert q4.months == ("2026-01", "2026-02") and q4.partial
    # VAT category A periods end in odd months (Jan, Mar...), B in even ones.
    assert [p.label for p in vat_periods(["2026-01", "2026-02", "2026-03"], "A")] == [
        "Dec 2025 - Jan 2026", "Feb 2026 - Mar 2026"]
    assert [p.label for p in vat_periods(["2026-01", "2026-02", "2026-03"], "B")] == [
        "Jan 2026 - Feb 2026", "Mar 2026 - Apr 2026"]
