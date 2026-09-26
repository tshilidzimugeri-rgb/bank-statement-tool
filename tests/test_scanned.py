"""Scanned statements: OCR clean-up, splitting a PDF into its statements, and
repairs of misread figures that are only kept when the statement's own totals
confirm them."""
from pathlib import Path

from statement_tool import config as config_mod
from statement_tool.extract.document import parse_statement_text, repair_from_balances, split_statements
from statement_tool.extract.ocr_extract import clean_ocr_text

PROJECT_ROOT = Path(__file__).parent.parent
LAYOUTS, GENERIC = config_mod.load_bank_layouts(PROJECT_ROOT / "config" / "banks.yaml")
FNB = LAYOUTS["fnb"]


def test_clean_ocr_text_removes_scan_marks_but_keeps_amounts():
    raw = "(02 Dec |Rtc Credit Transfer 10,000.00Cr| 10,129.74Cr|\n17 Dec |FNB App Transfer 600,00 470.26 * 8.00)"
    assert clean_ocr_text(raw).split("\n") == [
        "02 Dec Rtc Credit Transfer 10,000.00Cr 10,129.74Cr",
        "17 Dec FNB App Transfer 600.00 470.26 8.00",
    ]
    assert clean_ocr_text("Platinum Business Account : 62000000000") == "Platinum Business Account : 62000000000"
    assert clean_ocr_text("Balance 2,299.24Cr' 8.00") == "Balance 2,299.24Cr 8.00"
    assert clean_ocr_text("1,472.60Cr) 10,000") == "1,472.60Cr 10,000"  # thousands untouched


def test_pdf_split_into_statements_by_period():
    pages = [
        "Statement Period : 1 January 2026 to 31 January 2026\nrows",
        "Transactions continued",
        "Statement Period : 1 February 2026 to 28 February 2026\nrows",
        "Statement Period : 1 February 2026 to 28 February 2026\nsame statement, header repeated",
    ]
    assert split_statements(pages, LAYOUTS, GENERIC) == [[0, 1], [2, 3]]


STATEMENT = """First National Bank fnb.co.za
Business Account : 62000000000
Statement Period : 1 March 2026 to 31 March 2026
Opening Balance 1,000.00Cr
02 Mar Rtc Credit Transfer 500.00Cr 1,500.00Cr
05 Mar Payment To Supplier 300.00 1,200.00Cr
09 Mar Payment To Wages 700.00 500.00Cr
20 Mar #Monthly Account Fee 100.00 400.00Cr
Closing Balance 400.00Cr
No. Credit Transactions 1 500.00Cr
No. Debit Transactions 3 1,100.00Dr"""


def _read(text):
    return parse_statement_text(text, "scan.pdf", layouts=LAYOUTS, generic=GENERIC, client_rules=[],
                                sender=None, subject=None, interactive=False, used_ocr=True)


def test_clean_scan_reads_without_repairs():
    result = _read(STATEMENT)
    assert result.problems == [] and not (result.warning or "").startswith("scanned statement")
    assert result.account_number == "62000000000"
    assert [(t.date, t.debit, t.credit) for t in result.transactions][:2] == [
        ("2026-03-02", None, 500.0), ("2026-03-05", 300.0, None)]


def test_misread_amount_repaired_and_confirmed_by_totals():
    result = _read(STATEMENT.replace("Payment To Supplier 300.00", "Payment To Supplier 800.00"))
    assert result.problems == []
    assert result.transactions[1].debit == 300.0
    assert "amount read as 800.00, is 300.00" in result.warning


def test_misread_balance_repaired():
    result = _read(STATEMENT.replace("700.00 500.00Cr", "700.00 900.00Cr"))
    assert result.problems == []
    assert result.transactions[2].balance == 500.0


def test_two_misreads_on_one_row_are_not_guessed():
    result = _read(STATEMENT.replace("300.00 1,200.00Cr", "800.00 1,900.00Cr"))
    assert result.problems  # refused, not "repaired"


def test_no_repairs_without_printed_totals_to_confirm_them():
    no_totals = "\n".join(l for l in STATEMENT.split("\n") if not l.startswith("No."))
    result = _read(no_totals.replace("Payment To Supplier 300.00", "Payment To Supplier 800.00"))
    assert result.problems


def test_repair_that_breaks_the_printed_totals_is_rejected():
    rows = [("2026-03-02", "A", None, 500.0, 1500.0), ("2026-03-05", "B", 999.0, None, 1200.0)]
    text = "Closing Balance 1,200.00Cr\nNo. Credit Transactions 1 400.00Cr\nNo. Debit Transactions 1 300.00Dr"
    repaired, notes = repair_from_balances(rows, 1000.0, text, FNB, GENERIC)
    assert repaired is None and notes == []  # 999 -> 300 fits the balances, but money in doesn't match