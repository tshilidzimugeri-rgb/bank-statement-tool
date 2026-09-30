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
        Paragraph("Opening Balance: 1,000.00", styles["Normal"]),
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
    ("09/03/2026", "Groceries", "-300.00", "", "4,500.00"),  # this statement prints every debit negative
]


def test_clean_statement_has_no_problems_and_negative_debit_stays_a_debit(tmp_path):
    result = _parse(_statement_pdf(tmp_path / "s.pdf", GOOD_ROWS, footer="Closing Balance: 4,500.00"))
    assert result.ok and result.problems == []
    assert [(t.debit, t.credit) for t in result.transactions] == [(None, 5000.0), (1200.0, None), (300.0, None)]


def test_row_that_doesnt_add_up_is_added_and_marked_for_checking(tmp_path):
    rows = [*GOOD_ROWS[:2], ("09/03/2026", "Groceries", "-30.00", "", "4,500.00")]  # 300 misread as 30
    result = _parse(_statement_pdf(tmp_path / "s.pdf", rows))
    assert result.ok and len(result.transactions) == 3  # nothing turned away
    groceries = result.transactions[2]
    assert groceries.check  # but highlighted for a person to check
    assert [t.check for t in result.transactions[:2]] == ["", ""]  # the rows that add up aren't


def test_missing_last_row_caught_by_printed_closing_balance(tmp_path):
    result = _parse(_statement_pdf(tmp_path / "s.pdf", GOOD_ROWS[:2], footer="Closing Balance: 4,500.00"))
    assert any("closing balance is 4,500.00 but the last balance read is 4,800.00" in p for p in result.problems)
    assert result.status == "REVIEW_REQUIRED"


def test_unreadable_date_is_left_empty_and_marked(tmp_path):
    rows = [*GOOD_ROWS[:2], ("31/02/2026", "Groceries", "-300.00", "", "4,500.00")]  # no 31 February
    result = _parse(_statement_pdf(tmp_path / "s.pdf", rows))
    groceries = result.transactions[2]
    assert groceries.date == "" and "date unreadable" in groceries.check  # never borrowed from another row
    assert groceries.status == "REVIEW_REQUIRED"
    assert groceries.debit == 300.0  # the amount itself still adds up


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


def test_same_purchase_twice_on_different_printed_lines_is_not_a_duplicate():
    from statement_tool.extract.reading import Row, _flag_duplicates

    def flag(row, reason):
        row.check = reason

    # A payment's balance is printed after its fee, so the payment row has
    # none of its own; two such payments on different lines are two payments.
    first = Row("2026-07-01", "Online Purchase: Streamco", 25.00, None, None,
                evidence="01/07/2026 Online Purchase: Streamco -25.00 -2.00 150.00")
    second = Row("2026-07-01", "Online Purchase: Streamco", 25.00, None, None,
                 evidence="01/07/2026 Online Purchase: Streamco -25.00 -2.00 123.00")
    read_twice = Row("2026-07-01", "Online Purchase: Streamco", 25.00, None, None,
                     evidence="01/07/2026 Online Purchase: Streamco -25.00 -2.00 123.00")
    # Two R2.00 fees on different lines, the same balance after both (money
    # came in between their payments): two fees.
    fee_a = Row("2026-07-19", "Fee: PayShap", 2.0, None, 30.00, evidence="19/07/2026 PayShap -300.00 -2.00 30.00")
    fee_b = Row("2026-07-19", "Fee: PayShap", 2.0, None, 30.00, evidence="19/07/2026 PayShap -120.00 -2.00 30.00")
    _flag_duplicates([first, second, read_twice, fee_a, fee_b], flag)
    assert (first.check, second.check, fee_a.check, fee_b.check) == ("", "", "", "")
    assert read_twice.check == "possible duplicate of an earlier row"  # the same printed line again


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


def test_downloaded_workbook_restored_under_its_account_name(tmp_path):
    from statement_tool.excel_writer import restore_workbook

    original = tmp_path / "made" / "FTW Properties (10237421516).xlsx"
    t = _t("2026-03-01", "RENT", credit=100.0, balance=100.0)
    t.client, t.account = "FTW Properties", "10237421516"
    append_transactions(original, [t])

    work = tmp_path / "online"
    # Browsers rename a repeat download - the account comes from the rows.
    restored = restore_workbook(original.read_bytes(), work, "FTW Properties (10237421516) (1).xlsx")
    assert restored == work / "FTW Properties (10237421516).xlsx"
    assert len(read_transactions(restored)) == 1
    assert list(work.glob("~*")) == []

    assert restore_workbook(b"not a workbook", work, "junk.xlsx") is None
    assert list(work.glob("~*")) == []


def test_reading_a_workbook_does_not_block_the_next_save(tmp_path):
    book = tmp_path / "out.xlsx"
    append_transactions(book, [_t("2026-03-01", "A", credit=1.0, balance=1.0)])
    read_transactions(book)  # e.g. the page showing the overview
    assert append_transactions(book, [_t("2026-03-02", "B", credit=1.0, balance=2.0)]).added == 1


def test_statement_from_a_bank_with_no_settings_is_read_and_checked(tmp_path):
    # A bank not in banks.yaml, printing FNB-style rows: no year on dates,
    # attached Cr, an accrued-charges column and a no-description fee row.
    from reportlab.pdfgen import canvas

    lines = [
        "Mystery Bank Limited",
        "Account : 55512345678",
        "Statement Period : 1 December 2026 to 5 January 2027",
        "Opening Balance 1,000.00Cr",
        "Date Description Amount Balance Accrued Charges",
        "28 Dec Salary 5,000.00Cr 6,000.00Cr",
        "30 Dec Card Purchase Spar 250.00 5,750.00Cr 3.00",
        "03 Jan 3.00 5,747.00Cr",
        "Closing Balance 5,747.00Cr",
    ]
    path = tmp_path / "mystery.pdf"
    c = canvas.Canvas(str(path))
    for i, line in enumerate(lines):
        c.drawString(40, 800 - 18 * i, line)
    c.save()

    result = _parse(path)
    assert result.problems == []
    assert result.account_number == "55512345678"
    assert [(t.date, t.debit, t.credit, t.balance) for t in result.transactions] == [
        ("2026-12-28", None, 5000.0, 6000.0),  # December gets the year before the period's end
        ("2026-12-30", 250.0, None, 5750.0),
        ("2027-01-03", 3.0, None, 5747.0),
    ]
