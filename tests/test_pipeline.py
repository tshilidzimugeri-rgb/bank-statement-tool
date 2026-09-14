"""End-to-end tests against the synthetic fixtures in tests/fixtures/.

Run `python tests/fixtures/generate_fixtures.py` first if the PDFs are
missing (they're generated, not committed as binary fixtures).
"""
from pathlib import Path

import pytest

from statement_tool import config as config_mod
from statement_tool.extract.parser import parse_statement

FIXTURES = Path(__file__).parent / "fixtures"
PROJECT_ROOT = Path(__file__).parent.parent


def _settings():
    return config_mod.load_settings(dotenv_path=PROJECT_ROOT / ".env.example")


def _layouts():
    return config_mod.load_bank_layouts(PROJECT_ROOT / "config" / "banks.yaml")


def _client_rules():
    return config_mod.load_client_rules(PROJECT_ROOT / "config" / "clients.yaml")


@pytest.fixture(scope="module")
def fnb_pdf():
    path = FIXTURES / "sample_fnb_statement.pdf"
    if not path.exists():
        pytest.skip("run tests/fixtures/generate_fixtures.py first")
    return path


@pytest.fixture(scope="module")
def capitec_pdf():
    path = FIXTURES / "sample_capitec_statement.pdf"
    if not path.exists():
        pytest.skip("run tests/fixtures/generate_fixtures.py first")
    return path


def test_fnb_split_columns_parsed_correctly(fnb_pdf):
    layouts, generic = _layouts()
    result = parse_statement(
        fnb_pdf,
        layouts=layouts,
        generic=generic,
        client_rules=_client_rules(),
        settings=_settings(),
        interactive=False,
    )
    assert result.ok
    assert not result.used_ocr
    assert result.bank_display_name == "FNB"
    assert result.client == "Acme Trading CC"
    assert result.statement_period == "01 Jan 2024 to 31 Jan 2024"

    # Opening/closing balance rows have no debit or credit and must be
    # excluded, not turned into phantom zero-value transactions.
    descriptions = [t.description for t in result.transactions]
    assert "Opening Balance" not in descriptions
    assert len(result.transactions) == 7

    salary = next(t for t in result.transactions if t.description == "Salary Payment")
    assert salary.credit == 25000.0
    assert salary.debit is None
    assert salary.balance == 34549.25
    assert salary.date == "2024-01-05"


def test_capitec_signed_amount_column_split_into_debit_credit(capitec_pdf):
    layouts, generic = _layouts()
    result = parse_statement(
        capitec_pdf,
        layouts=layouts,
        generic=generic,
        client_rules=_client_rules(),
        settings=_settings(),
        interactive=False,
    )
    assert result.ok
    assert result.bank_display_name == "Capitec"
    assert result.client == "Jane Dlamini"

    salary = next(t for t in result.transactions if t.description == "Salary")
    assert salary.credit == 18500.0
    assert salary.debit is None

    groceries = next(t for t in result.transactions if "Grocery" in t.description)
    assert groceries.debit == 620.40
    assert groceries.credit is None


def test_unrecognized_client_is_flagged_not_silently_dropped():
    layouts, generic = _layouts()
    path = FIXTURES / "sample_standardbank_statement_unmapped.pdf"
    if not path.exists():
        pytest.skip("run tests/fixtures/generate_fixtures.py first")
    result = parse_statement(
        path,
        layouts=layouts,
        generic=generic,
        client_rules=_client_rules(),
        settings=_settings(),
        interactive=False,
    )
    assert result.ok
    assert result.client == "UNMAPPED_CLIENT"
    assert result.warning and "could not be matched" in result.warning


def test_corrupt_pdf_is_reported_as_failure_not_a_crash(tmp_path):
    layouts, generic = _layouts()
    bad_file = tmp_path / "corrupt.pdf"
    bad_file.write_text("not a real pdf")
    result = parse_statement(
        bad_file,
        layouts=layouts,
        generic=generic,
        client_rules=_client_rules(),
        settings=_settings(),
        interactive=False,
    )
    assert result.ok is False
    assert result.error
