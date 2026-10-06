"""Categories from config/categories.yaml: general South African rules that
know money in from money out, and VAT that follows the business's VAT
registration. Made-up descriptions in the banks' own wording."""
from pathlib import Path

import pytest

from statement_tool.categorize import Category, categorize, for_vat_registration, load_categories
from statement_tool.excel_writer import append_transactions, read_transactions, vat_by_month
from statement_tool.models import Transaction

CATEGORIES = load_categories(Path(__file__).parent.parent / "config" / "categories.yaml")


@pytest.mark.parametrize("description, money_in, expected", [
    ("Fee: Banking App Immediate Payment: J Smith", False, "Bank charges"),
    ("#Item Unpaid No Funds 03", False, "Bank charges"),
    ("Magtape Unpaid Not Provided For", True, "Refunds & reversals"),
    ("Live Better Interest Sweep", False, "Savings & own-account transfers"),
    ("FNB App Transfer From Savings", True, "Savings & own-account transfers"),
    ("Interest Received", True, "Interest received"),
    ("Interest On Debit Balance", False, "Interest paid"),
    ("Adjustment Dr SARS VAT", False, "Tax payments (SARS)"),
    ("Debit Order WesBank_Fi123", False, "Loan & finance repayments"),
    ("ATM Cash Withdrawal: Ncr Pretoria", False, "Cash withdrawals & sends"),
    ("Payshap Account Off-Us Teacher Salary", False, "Salaries & wages"),
    ("Dis-Chem Polokwane (Card 1234) Pharmacy", False, "Medical & health"),
    ("Debit Kingprice Ref123", False, "Insurance"),
    ("FNB App Prepaid Airtime 0820000000", False, "Telephone, airtime & data"),
    ("VAS0012 | Electricity Purchase", False, "Water & electricity"),
    ("Sasol Theresa 1234*5678 10 May", False, "Fuel"),
    ("DL UBER EATS 1234*5678", False, "Meals & takeaways"),
    ("DL*BOLT 1234*5678 22 Aug", False, "Transport & travel"),
    ("Shoprite Sb047923 (Card 1234) Groceries", False, "Groceries & supplies"),
    ("Apple.com/Bill Cork", False, "Software & subscriptions"),
    ("Pep Stores (Card 1234) Clothing & Shoes", False, "Equipment, clothing & retail"),
    ("Banking App Immediate Payment: J Smith", False, "Payments to others"),
    ("Payment Received: J Smith", True, "Payments received"),
    ("ADT Cash Deposit Megacity", True, "Payments received"),
    # No accidental hits.
    ("Coffee Shop Rosebank", False, "Meals & takeaways"),
    ("Current Account Something", False, "Other expenses (uncategorised)"),
])
def test_south_african_descriptions(description, money_in, expected):
    assert categorize(description, money_in, CATEGORIES) == expected


def test_rules_know_money_in_from_money_out():
    # A deposit mentioning "municipal" isn't a (negative) electricity expense.
    assert categorize("Municipal Deposit Refund", True, CATEGORIES) == "Refunds & reversals"
    assert categorize("Municipal Account Payment", True, CATEGORIES) != "Water & electricity"
    assert categorize("Municipality of Polokwane", False, CATEGORIES) == "Water & electricity"
    # Categories made in code (without a direction) still match both ways.
    assert categorize("RENT", True, [Category("Rent", "income", ["rent"])]) == "Rent"


def _t(day, desc, debit=None, credit=None, balance=None):
    return Transaction(client="", bank="Test Bank", statement_period="P1", date=day, description=desc,
                       debit=debit, credit=credit, balance=balance, source_file="a.pdf", account="1234567890")


def test_vat_follows_the_business_vat_registration(tmp_path):
    rows = [_t("2026-03-02", "Payment Received: J Smith", credit=1150.0, balance=1150.0),
            _t("2026-03-03", "FNB App Prepaid Airtime 0820000000", debit=115.0, balance=1035.0),
            _t("2026-03-04", "Sasol Theresa 1234*5678", debit=500.0, balance=535.0)]
    registered = tmp_path / "registered.xlsx"
    append_transactions(registered, rows, for_vat_registration(CATEGORIES, True))
    vat = vat_by_month(read_transactions(registered), CATEGORIES)[0]
    assert round(vat["VAT on income"], 2) == 150.00  # 1 150 x 15/115
    assert round(vat["VAT on expenses"], 2) == 15.00  # airtime; fuel is zero-rated
    not_registered = tmp_path / "not.xlsx"
    append_transactions(not_registered, rows, for_vat_registration(CATEGORIES, False))
    vat = vat_by_month(read_transactions(not_registered), CATEGORIES)[0]
    assert (vat["VAT on income"], vat["VAT on expenses"]) == (0.0, 0.0)
    # Switching registration on later works the VAT out afresh for every row.
    append_transactions(not_registered, [], for_vat_registration(CATEGORIES, True), reset_vat=True)
    vat = vat_by_month(read_transactions(not_registered), CATEGORIES)[0]
    assert round(vat["VAT on income"], 2) == 150.00
