from statement_tool.extract.amounts import parse_amount
from statement_tool.extract.dates import parse_date


def test_parse_amount_basic():
    assert parse_amount("1,234.56") == 1234.56
    assert parse_amount("R 450.75") == 450.75
    assert parse_amount("-620.40") == -620.40


def test_parse_amount_parentheses_and_dr_are_negative():
    assert parse_amount("(1,234.56)") == -1234.56
    assert parse_amount("1234.56 Dr") == -1234.56


def test_parse_amount_blank_placeholders_are_none():
    assert parse_amount("") is None
    assert parse_amount("-") is None
    assert parse_amount(None) is None


def test_parse_amount_thousands_and_decimal_comma_styles():
    assert parse_amount("1.234,56") == 1234.56
    assert parse_amount("1234,56") == 1234.56


def test_parse_date_common_formats():
    assert parse_date("03/01/2024") == "2024-01-03"
    assert parse_date("2024-01-03") == "2024-01-03"
    assert parse_date("03 Jan 2024") == "2024-01-03"


def test_parse_date_unparsable_returns_none():
    assert parse_date("not a date") is None
    assert parse_date("") is None
