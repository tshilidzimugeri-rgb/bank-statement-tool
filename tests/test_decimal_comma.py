"""Statements that write figures with a decimal comma ("1 234,56",
"13,50-"), and entries printed over two lines with the date once: the
dated line without a balance, the next line with a fee and the balance.
Made-up values."""
from datetime import date
from pathlib import Path

from statement_tool import config as config_mod
from statement_tool.extract import auto_extract
from statement_tool.extract.bank_detect import extract_statement_period
from statement_tool.extract.reading import auto_reading

PROJECT_ROOT = Path(__file__).parent.parent
_, GENERIC = config_mod.load_bank_layouts(PROJECT_ROOT / "config" / "banks.yaml")
PERIOD = (date(2026, 3, 1), date(2026, 3, 31))

TEXT = """\
Karoo Savings Bank
Account number 12 3456 7890
1 Mar 2026 to 31 Mar 2026
Balance brought forward 1 000,00
Total credits 2 500,00
Total debits 1 300,00-
Total service fees (R1,30 VAT included) 10,00
date description transaction amount R balance R
1 Mar 2026 Balance Brought Forward 1 000,00
2 Mar 2026 Salary Credit 2 500,00 3 500,00
3 Mar 2026 Instant Payment Debit 1 200,00-
Service Fee 8,75- 2 291,25
Ref: Rent March
4 Mar 2026 Payment Notification
Service Fee 1,25- 2 290,00
5 Mar 2026 Card Purchase 100,00- 2 190,00
Please turn over
"""


def test_decimal_comma_statement_read_exactly_with_fees_on_their_own_lines():
    reading = auto_reading(TEXT, PERIOD, TEXT, GENERIC, GENERIC)

    assert [(r.date, r.debit, r.credit, r.balance) for r in reading.rows] == [
        ("2026-03-02", None, 2500.00, 3500.00),
        # The balance between the payment and its fee isn't printed, so none is made up.
        ("2026-03-03", 1200.00, None, None),
        ("2026-03-03", 8.75, None, 2291.25),
        ("2026-03-04", 1.25, None, 2290.00),
        ("2026-03-05", 100.00, None, 2190.00),
    ]
    assert [r.check for r in reading.rows] == [""] * 5
    # Fees totalled apart from the debits: 1 300,00 + 10,00 is what was read.
    assert reading.problems == []
    assert (reading.report["printed_total_debits"], reading.report["printed_total_fees"]) == (1300.00, 10.00)
    assert reading.rows[2].description == "Instant Payment Debit | Service Fee | Ref: Rent March"
    assert reading.rows[3].description == "Payment Notification | Service Fee"
    # The evidence shows the dated line the fee's date comes from.
    assert reading.rows[3].evidence == "4 Mar 2026 Payment Notification / Service Fee 1,25- 2 290,00"
    assert reading.rows[4].description == "Card Purchase"  # the page footer isn't part of it


def test_wrong_fee_total_is_not_approved():
    text = TEXT.replace("VAT included) 10,00", "VAT included) 11,00")
    reading = auto_reading(text, PERIOD, text, GENERIC, GENERIC)
    assert any("money out" in p for p in reading.problems)
    assert all(r.check for r in reading.rows)


def test_period_printed_on_a_line_of_its_own():
    assert extract_statement_period(TEXT, GENERIC, GENERIC) == "1 Mar 2026 to 31 Mar 2026"


def test_undated_line_after_a_complete_entry_is_not_given_a_date():
    # The dated line above already has its balance, so nothing says this
    # line belongs to it: it keeps no date and is marked.
    text = ("Opening balance 100,00\n1 Mar 2026 Card Purchase 10,00- 90,00\nTransfer 5,00- 85,00\n"
            "2 Mar 2026 Card Purchase 5,00- 80,00\n")
    rows = auto_extract.extract(text, PERIOD[1], PERIOD[0]).rows
    assert [(r.date_iso, r.debit, r.balance) for r in rows] == [
        ("2026-03-01", 10.00, 90.00), (None, 5.00, 85.00), ("2026-03-02", 5.00, 80.00)]
    assert "no date printed" in rows[1].check


def test_each_statement_is_read_in_its_own_decimal_style():
    assert auto_extract.decimal_mark("01 Mar PAYMENT 1,234.56 12.50") == "."
    assert auto_extract.decimal_mark("01 Mar PAYMENT 1 234,56- 12,50") == ","
    # "1 234,56" may be one number, as "1 234.56" may.
    tokens = auto_extract.money_tokens("5 Mar Card 1 234,56- 12 345,67", ",")
    assert [[(m.value, m.sign) for m in t] for t in tokens] == [
        [(234.56, -1), (1234.56, -1)], [(345.67, None), (12345.67, None)]]
    assert [t[0].value for t in auto_extract.money_tokens("5 Mar Card 1.234,56 12,50-", ",")] == [1234.56, 12.50]
    # In a statement written "1,234.56", "12,50" is not an amount.
    assert [t[0].value for t in auto_extract.money_tokens("REF 12,50 100.00 1,000.00")] == [100.00, 1000.00]
