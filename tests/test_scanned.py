"""Scanned and multi-statement documents, and the extraction rules: printed
values are never changed or guessed, disagreements between the two
independent readings are flagged, and anything uncertain is
REVIEW_REQUIRED."""
from pathlib import Path

from statement_tool import config as config_mod
from statement_tool.extract.document import parse_statement_text, split_statements
from statement_tool.extract.ocr_extract import clean_ocr_text

PROJECT_ROOT = Path(__file__).parent.parent
LAYOUTS, GENERIC = config_mod.load_bank_layouts(PROJECT_ROOT / "config" / "banks.yaml")


def test_clean_ocr_text_removes_scan_marks_but_never_changes_amounts():
    raw = "(02 Dec |Rtc Credit Transfer 10,000.00Cr| 10,129.74Cr|\n17 Dec |FNB App Transfer 600,00 470.26 * 8.00)"
    assert clean_ocr_text(raw).split("\n") == [
        "02 Dec Rtc Credit Transfer 10,000.00Cr 10,129.74Cr",
        "17 Dec FNB App Transfer 600,00 470.26 8.00",  # decimal comma left as read
    ]
    assert clean_ocr_text("Platinum Business Account : 62000000000") == "Platinum Business Account : 62000000000"
    assert clean_ocr_text("Balance 2,299.24Cr' 8.00") == "Balance 2,299.24Cr 8.00"


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


def _read(text, second=STATEMENT):
    return parse_statement_text(text, "scan.pdf", second_text=second, layouts=LAYOUTS, generic=GENERIC,
                                client_rules=[], sender=None, subject=None, interactive=False, used_ocr=True)


def test_clean_scan_read_twice_the_same_is_approved():
    result = _read(STATEMENT)
    assert result.status == "APPROVED" and result.problems == []
    assert {t.status for t in result.transactions} == {"APPROVED"}
    assert {t.confidence for t in result.transactions} == {0.97}  # a scan: verified, but not quite 1.00
    assert result.transactions[1].evidence == "05 Mar Payment To Supplier 300.00 1,200.00Cr"
    assert result.report["computed_closing"] == 400.0 == result.report["printed_closing"]


def test_misread_amount_is_kept_as_printed_and_flagged():
    misread = STATEMENT.replace("Payment To Supplier 300.00", "Payment To Supplier 800.00")
    result = _read(misread, second=misread)
    supplier = result.transactions[1]
    assert supplier.status == "REVIEW_REQUIRED"
    assert supplier.debit is None or supplier.debit == 800.0  # never replaced by the 300 the balances imply
    assert "printed as 800.00 but the balances give 300.00" in supplier.check
    assert result.status == "REVIEW_REQUIRED"


def test_rows_the_two_readings_disagree_on_are_flagged():
    other = STATEMENT.replace("700.00 500.00Cr", "780.00 420.00Cr").replace("100.00 400.00Cr", "100.00 320.00Cr")
    result = _read(STATEMENT, second=other)
    flagged = [t.description for t in result.transactions if t.status == "REVIEW_REQUIRED"]
    assert flagged == ["Payment To Wages", "#Monthly Account Fee"]
    assert "other OCR pass" in result.transactions[2].check


def test_without_a_second_reading_nothing_is_approved():
    result = _read(STATEMENT, second=None)
    assert {t.status for t in result.transactions} == {"REVIEW_REQUIRED"}


def test_unreadable_date_is_left_empty_not_borrowed():
    garbled = STATEMENT.replace("09 Mar Payment", "41 Mar Payment")
    result = _read(garbled, second=garbled)
    wages = result.transactions[2]
    assert wages.date == "" and "date unreadable" in wages.check and wages.status == "REVIEW_REQUIRED"
    assert wages.debit == 700.0  # the amount itself is confirmed by the balances


def test_direction_not_established_is_left_unassigned():
    no_opening = "\n".join(l for l in STATEMENT.split("\n") if not l.startswith("Opening"))
    no_opening = no_opening.replace("500.00Cr 1,500.00Cr", "500.00 1,500.00Cr")  # first amount unmarked
    result = _read(no_opening, second=no_opening)
    first = result.transactions[0]
    assert first.unassigned == 500.0 and first.debit is None and first.credit is None
    assert first.status == "REVIEW_REQUIRED"