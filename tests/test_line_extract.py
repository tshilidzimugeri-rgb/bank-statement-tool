"""Line-based fallback extraction and encrypted-PDF handling."""
from pathlib import Path

import pytest

from statement_tool import config as config_mod
from statement_tool.extract import line_extract
from statement_tool.extract.parser import PdfPasswordError, _find_password, parse_statement

PROJECT_ROOT = Path(__file__).parent.parent

# Shape of Standard Bank's 6-month statement text layer (synthetic values).
STANDARD_BANK_6M_TEXT = """\
Transaction details Available Balance:R19,050.00
Date Description Payments Deposits Balance
STATEMENT OPENING BALANCE 1,000.00
28 Feb 26 CARTRACK S105778598 -97.68 902.32
DEBIT TRANSFER
28 Feb 26 FEE - DEBIT ORDER -4.00 898.32
FEE - DEBIT ORDER
28 Feb 26 0000010237421516 00009 R4.05 -4.05 894.27
FEE: MYUPDATES FOR BUSINESS
The Standard Bank of South Africa Limited (Reg. No. 1962/000738/06. Authorised financial service provider.
03 Mar 26 CHAMELA 63139171598 18,155.73 19,050.00
REAL TIME TRANSFER FROM
"""


def test_signed_lines_split_into_debit_and_credit_with_continuation():
    result = line_extract.extract(STANDARD_BANK_6M_TEXT)

    assert result.unreconciled == 0
    assert [(r.date_raw, r.debit, r.credit, r.balance) for r in result.rows] == [
        ("28 Feb 26", 97.68, None, 902.32),
        ("28 Feb 26", 4.00, None, 898.32),
        ("28 Feb 26", 4.05, None, 894.27),
        ("03 Mar 26", None, 18155.73, 19050.00),
    ]
    assert result.rows[0].description == "CARTRACK S105778598 | DEBIT TRANSFER"
    # Continuation identical to the description isn't repeated.
    assert result.rows[1].description == "FEE - DEBIT ORDER"
    # An "R4.05" inside the description isn't taken as the amount.
    assert result.rows[2].description == "0000010237421516 00009 R4.05 | FEE: MYUPDATES FOR BUSINESS"


def test_long_footer_line_not_appended_as_continuation():
    text = "01 Mar 26 SOMETHING -1.00 99.00\n" + "X" * 100 + "\n"
    result = line_extract.extract(text)
    assert result.rows[0].description == "SOMETHING"


def test_unsigned_amounts_resolved_from_running_balance():
    text = (
        "Opening balance 100.00\n"
        "01 Mar 26 GROCERIES 30.00 70.00\n"
        "02 Mar 26 SALARY 500.00 570.00\n"
    )
    result = line_extract.extract(text)
    assert [(r.debit, r.credit) for r in result.rows] == [(30.00, None), (None, 500.00)]
    assert result.unreconciled == 0


def test_unreconciled_row_is_counted():
    text = "Opening balance 100.00\n01 Mar 26 ODD ONE -30.00 50.00\n"
    result = line_extract.extract(text)
    assert result.unreconciled == 1
    assert result.rows[0].debit == 30.00


@pytest.fixture
def encrypted_pdf(tmp_path):
    from reportlab.pdfgen import canvas

    path = tmp_path / "locked.pdf"
    c = canvas.Canvas(str(path), encrypt="s3cret")
    c.drawString(72, 720, "Standard Bank statement")
    c.drawString(72, 700, "Opening balance 100.00")
    c.drawString(72, 680, "01 Mar 26 GROCERIES -30.00 70.00")
    c.save()
    return path


def test_find_password_tries_configured_passwords(encrypted_pdf):
    assert _find_password(encrypted_pdf, ["wrong", "s3cret"]) == "s3cret"
    with pytest.raises(PdfPasswordError):
        _find_password(encrypted_pdf, ["wrong"])


def test_encrypted_statement_reports_missing_password_then_parses_with_it(encrypted_pdf):
    settings = config_mod.load_settings(dotenv_path=PROJECT_ROOT / ".env.example")
    layouts, generic = config_mod.load_bank_layouts(PROJECT_ROOT / "config" / "banks.yaml")
    kwargs = dict(layouts=layouts, generic=generic, client_rules=[], interactive=False)

    settings.pdf_passwords = []
    locked = parse_statement(encrypted_pdf, settings=settings, **kwargs)
    assert not locked.ok
    assert "STATEMENT_PDF_PASSWORDS" in locked.error

    settings.pdf_passwords = ["s3cret"]
    opened = parse_statement(encrypted_pdf, settings=settings, **kwargs)
    assert opened.ok, opened.error
    assert [(t.date, t.debit) for t in opened.transactions] == [("2026-03-01", 30.00)]


# Shape of Capitec's Transaction History (synthetic values): signed amounts
# with space thousands, an optional fee before the balance, VAT asterisks,
# and summary lists above/below that are not transactions.
CAPITEC_TEXT = """\
From Date: 01/08/2026 Opening Balance: R1 000.00
To Date: 31/08/2026 Closing Balance: R3 087.00
Money In Summary R5 000.00 Money Out Summary -R2 913.00
Card Subscriptions
02/08/2026 Uber -R58.00
Date Description Category Money In Money Out Fee* Balance
01/08/2026 Payment Received: Salary Other Income 5 000.00 6 000.00
02/08/2026 Banking App External PayShap Payment: J Doe (064 014 Digital Payments -2 900.00 -2.00 3 098.00
4653)
03/08/2026 Banking App Prepaid Purchase: Electricity Electricity -11.00* 3 087.00
* Includes VAT at 15%
Pending Card Transactions
04/08/2026 Uber Cpt (Card 1234) -R53.22
"""


def test_capitec_rows_with_fee_space_thousands_and_vat_marker():
    result = line_extract.extract(CAPITEC_TEXT, thousands=" ", fee_column=True)

    assert result.opening_balance == 1000.00
    assert result.skipped_lines == []  # the Uber summary/pending lines aren't transactions
    assert [(r.description, r.debit, r.credit, r.balance) for r in result.rows] == [
        ("Payment Received: Salary Other Income", None, 5000.00, 6000.00),
        ("Banking App External PayShap Payment: J Doe (064 014 Digital Payments | 4653)", 2900.00, None, 3100.00),
        ("Fee: Banking App External PayShap Payment: J Doe (064 014 Digital Payments | 4653)", 2.00, None, 3098.00),
        ("Banking App Prepaid Purchase: Electricity Electricity", 11.00, None, 3087.00),
    ]


def test_capitec_statement_checks_pass_and_catch_a_dropped_fee():
    from statement_tool.extract.parser import _all_problems
    from statement_tool.extract.dates import parse_date

    layouts, generic = config_mod.load_bank_layouts(PROJECT_ROOT / "config" / "banks.yaml")

    def problems(text):
        r = line_extract.extract(text, thousands=" ", fee_column=True)
        raw = [(parse_date(x.date_raw), x.description, x.debit, x.credit, x.balance) for x in r.rows]
        return _all_problems(raw, r.opening_balance, text, layouts["capitec"], generic)

    assert problems(CAPITEC_TEXT) == []
    assert problems(CAPITEC_TEXT.replace("-2 900.00 -2.00 3 098.00", "-2 900.00 3 098.00"))


def test_space_thousands_not_used_for_comma_banks():
    # "Unit 141 200.00" must not become 141 200.00 for a comma-thousands bank.
    result = line_extract.extract("01 Mar 26 UNIT 141 200.00 1,200.00\n")
    assert [(r.description, r.credit) for r in result.rows] == [("UNIT 141", 200.00)]
