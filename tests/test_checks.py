"""Safeguards: statements whose numbers don't check out are caught, and the
workbook can't be corrupted or silently lose/merge transactions.
"""
from pathlib import Path

import pytest
from openpyxl import load_workbook
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table
from reportlab.lib.styles import getSampleStyleSheet

from statement_tool import config as config_mod
from statement_tool import excel_writer
from statement_tool.checks import balance_breaks, workbook_gaps
from statement_tool.excel_writer import WorkbookLockedError, append_transactions, read_transactions
from statement_tool.extract.parser import parse_statement
from statement_tool.models import Transaction

PROJECT_ROOT = Path(__file__).parent.parent


# --- balance chain ---------------------------------------------------------------

def test_balance_breaks_flags_rows_that_dont_follow():
    rows = [(None, 100.0, 1100.0), (50.0, None, 1050.0), (50.0, None, 900.0), (None, 10.0, 910.0)]
    breaks = balance_breaks(rows, opening=1000.0)
    assert [(b.index, b.expected, b.actual) for b in breaks] == [(2, 1000.0, 900.0)]


def test_balance_breaks_checks_first_row_against_opening():
    assert [b.index for b in balance_breaks([(None, 100.0, 1200.0)], opening=1000.0)] == [0]
    assert balance_breaks([(None, 100.0, 1200.0)]) == []  # no opening known: nothing to compare


def test_workbook_gap_found_when_a_months_statement_is_missing():
    march = [{"Client": "Dad", "Bank": "SB", "Date": "2026-03-05", "Description": "A",
              "Debit": 100.0, "Credit": None, "Balance": 900.0}]
    may = [{"Client": "Dad", "Bank": "SB", "Date": "2026-05-05", "Description": "B",
            "Debit": 100.0, "Credit": None, "Balance": 500.0}]  # April's R300 of spending is missing
    gaps = workbook_gaps(march + may)
    assert [(g.after_date, g.at_date, g.expected, g.actual) for g in gaps] == [("2026-03-05", "2026-05-05", 800.0, 500.0)]
    assert workbook_gaps(march) == []


# --- statement checks, end to end on generated PDFs ------------------------------------

def _statement_pdf(path: Path, rows: list[tuple], footer: str | None = None, account: bool = True) -> Path:
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(str(path), pagesize=A4)
    elements = [
        Paragraph("First National Bank", styles["Title"]),
        *([Paragraph("Account Number: 6205 1234 567", styles["Normal"])] if account else []),
        Paragraph("Statement Period: 01 Mar 2026 to 31 Mar 2026", styles["Normal"]),
        Table([["Date", "Description", "Debit", "Credit", "Balance"], *rows]),
    ]
    if footer:
        elements.append(Paragraph(footer, styles["Normal"]))
    doc.build(elements)
    return path


def _parse(path: Path):
    settings = config_mod.load_settings(dotenv_path=PROJECT_ROOT / ".env.example")
    layouts, generic = config_mod.load_bank_layouts(PROJECT_ROOT / "config" / "banks.yaml")
    return parse_statement(path, layouts=layouts, generic=generic, client_rules=[], settings=settings,
                           interactive=False)


GOOD_ROWS = [
    ("02/03/2026", "Salary", "", "5,000.00", "6,000.00"),
    ("05/03/2026", "Insurance", "-1,200.00", "", "4,800.00"),  # debit printed negative
    ("09/03/2026", "Groceries", "300.00", "", "4,500.00"),
]


def test_clean_statement_has_no_problems_and_negative_debit_stays_a_debit(tmp_path):
    result = _parse(_statement_pdf(tmp_path / "s.pdf", GOOD_ROWS, footer="Closing Balance: 4,500.00"))
    assert result.ok and result.problems == []
    assert [(t.debit, t.credit) for t in result.transactions] == [(None, 5000.0), (1200.0, None), (300.0, None)]


def test_row_that_doesnt_add_up_is_a_problem(tmp_path):
    rows = [*GOOD_ROWS[:2], ("09/03/2026", "Groceries", "30.00", "", "4,500.00")]  # 300 misread as 30
    result = _parse(_statement_pdf(tmp_path / "s.pdf", rows))
    assert any("don't add up with the running balance" in p for p in result.problems)


def test_missing_last_row_caught_by_printed_closing_balance(tmp_path):
    result = _parse(_statement_pdf(tmp_path / "s.pdf", GOOD_ROWS[:2], footer="Closing Balance: 4,500.00"))
    assert any("closing balance on the statement is 4,500.00" in p for p in result.problems)


def test_unreadable_date_is_a_problem(tmp_path):
    rows = [*GOOD_ROWS[:2], ("31/02/2026", "Groceries", "300.00", "", "4,500.00")]  # no 31 February
    result = _parse(_statement_pdf(tmp_path / "s.pdf", rows))
    assert any("date that couldn't be read" in p for p in result.problems)


def test_statement_without_balances_is_a_problem(tmp_path):
    rows = [(d, desc, debit, credit, "") for d, desc, debit, credit, _ in GOOD_ROWS]
    result = _parse(_statement_pdf(tmp_path / "s.pdf", rows))
    assert any("no running balance" in p for p in result.problems)


# --- workbook safety ---------------------------------------------------------------------

def _t(day, desc, debit=None, credit=None, balance=None):
    return Transaction(client="Dad", bank="SB", statement_period="P", date=day, description=desc,
                       debit=debit, credit=credit, balance=balance, source_file="a.pdf")


def test_workbook_open_in_excel_raises_clear_error_and_leaves_it_untouched(tmp_path, monkeypatch):
    book = tmp_path / "out.xlsx"
    append_transactions(book, [_t("2026-03-01", "FIRST", credit=1.0, balance=1.0)])
    before = book.read_bytes()

    def locked(*_args):
        raise PermissionError("file in use")

    monkeypatch.setattr(excel_writer.os, "replace", locked)
    with pytest.raises(WorkbookLockedError, match="close it in Excel"):
        append_transactions(book, [_t("2026-03-02", "SECOND", credit=1.0, balance=2.0)])
    assert book.read_bytes() == before
    assert not list(tmp_path.glob("~*"))  # no temp file left behind


def test_each_write_keeps_a_backup_of_the_previous_workbook(tmp_path, monkeypatch):
    monkeypatch.setattr(excel_writer, "BACKUPS_KEPT", 2)
    book = tmp_path / "out.xlsx"
    for day in range(1, 5):
        append_transactions(book, [_t(f"2026-03-0{day}", "X", credit=1.0, balance=float(day))])
    backups = sorted((tmp_path / "backups").glob("out-*.xlsx"))
    assert len(backups) == 2  # pruned to the newest BACKUPS_KEPT
    assert len(read_transactions(backups[-1])) == 3  # the state just before the last write


def test_identical_rows_without_balance_are_both_kept_but_not_re_added(tmp_path):
    book = tmp_path / "out.xlsx"
    twice = [_t("2026-03-01", "FEE", debit=4.0), _t("2026-03-01", "FEE", debit=4.0)]
    assert append_transactions(book, twice).added == 2
    again = append_transactions(book, twice)
    assert (again.added, again.skipped_duplicates) == (0, 2)
    assert len(read_transactions(book)) == 2


def test_merged_debit_credit_heading_is_not_read_as_one_column():
    from statement_tool.extract.column_map import map_header_row

    layouts, _ = config_mod.load_bank_layouts(PROJECT_ROOT / "config" / "banks.yaml")
    assert map_header_row(["Date", "Description", "Debit Credit", "Balance"], layouts["fnb"]) is None
    assert map_header_row(["Date", "Description", "Debit", "Credit", "Balance"], layouts["fnb"]) is not None


def test_workbook_that_cant_even_be_opened_gives_clear_error(tmp_path, monkeypatch):
    book = tmp_path / "out.xlsx"
    append_transactions(book, [_t("2026-03-01", "FIRST", credit=1.0, balance=1.0)])

    def locked(*_args, **_kwargs):
        raise PermissionError("file in use")

    monkeypatch.setattr(excel_writer, "load_workbook", locked)
    with pytest.raises(WorkbookLockedError, match="close it in Excel"):
        read_transactions(book)
    with pytest.raises(WorkbookLockedError, match="close it in Excel"):
        append_transactions(book, [_t("2026-03-02", "SECOND", credit=1.0, balance=2.0)])


# --- one workbook per account ------------------------------------------------------------

def test_second_account_is_refused_in_the_same_workbook(tmp_path):
    from statement_tool.excel_writer import MixedAccountsError

    book = tmp_path / "out.xlsx"
    mine = _t("2026-03-01", "SALARY", credit=100.0, balance=100.0)
    mine.account = "111"
    theirs = _t("2026-03-01", "SALARY", credit=100.0, balance=100.0)
    theirs.account = "222"
    append_transactions(book, [mine])
    with pytest.raises(MixedAccountsError):
        append_transactions(book, [theirs])
    assert len(read_transactions(book)) == 1


def test_each_account_gets_its_own_workbook_name(tmp_path):
    from statement_tool.excel_writer import workbook_path_for

    assert workbook_path_for(tmp_path, "FTW Properties", "10237421516", "x.pdf").name == "FTW Properties (10237421516).xlsx"
    assert workbook_path_for(tmp_path, "UNMAPPED_CLIENT", "1234567890", "x.pdf").name == "Account 1234567890.xlsx"
    # No account number: a workbook of its own per statement, never shared.
    assert workbook_path_for(tmp_path, "A/B", None, "may.pdf").name == "A_B - may.xlsx"
    assert workbook_path_for(tmp_path, None, None, "june.pdf").name == "Unknown account - june.xlsx"


@pytest.mark.parametrize(
    "bank, first_page, expected",
    [
        ("standard_bank", "Account number: 10 23 742 151 6 Address:", "10237421516"),
        ("capitec", "Tax Invoice\nAccount 1234567890 VAT Registration Number 4000000000", "1234567890"),
        ("fnb", "Account Number: 6205 1234 567", "62051234567"),  # generic fallback
    ],
)
def test_account_number_found_on_first_page(bank, first_page, expected):
    from statement_tool.extract.parser import _find_account_number

    layouts, generic = config_mod.load_bank_layouts(PROJECT_ROOT / "config" / "banks.yaml")
    assert _find_account_number(first_page, layouts[bank], generic) == expected


def test_identical_rows_sharing_a_balance_are_both_kept():
    # Two R2.00 fees on the same day with the same balance (money came in
    # between their payments) - both real, neither a duplicate.
    import tempfile
    book = Path(tempfile.mkdtemp()) / "out.xlsx"
    fee = lambda: _t("2026-09-01", "Fee: PayShap", debit=2.0, balance=1852.43)
    batch = [
        _t("2026-09-01", "PayShap", debit=136.0, balance=1854.43), fee(),
        _t("2026-09-01", "Received", credit=68.0, balance=1920.43),
        _t("2026-09-01", "PayShap", debit=66.0, balance=1854.43), fee(),
    ]
    assert balance_breaks([(t.debit, t.credit, t.balance) for t in batch]) == []
    assert append_transactions(book, batch).added == 5
    assert append_transactions(book, batch).added == 0  # re-adding the statement adds nothing


def test_statements_without_account_number_never_share_a_workbook(tmp_path):
    from statement_tool.excel_writer import MixedAccountsError

    book = tmp_path / "out.xlsx"
    may = _t("2026-05-01", "X", credit=1.0, balance=1.0)
    may.source_file = "may.pdf"
    june = _t("2026-06-01", "Y", credit=1.0, balance=2.0)
    june.source_file = "june.pdf"
    append_transactions(book, [may])
    assert append_transactions(book, [may]).added == 0  # the same statement again is fine
    with pytest.raises(MixedAccountsError):
        append_transactions(book, [june])


def test_missing_account_number_is_a_problem(tmp_path):
    # The generated statements have no account number on them.
    result = _parse(_statement_pdf(tmp_path / "s.pdf", GOOD_ROWS, account=False))
    assert any("no account number" in p for p in result.problems)
    assert not any("no account number" in p for p in _parse(_statement_pdf(tmp_path / "t.pdf", GOOD_ROWS)).problems)
