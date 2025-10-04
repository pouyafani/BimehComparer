# tests/test_utils.py
from src.bimeh_compare import parse_number, values_equal

def test_parse_number_basic():
    assert parse_number("1,234") == 1234.0

def test_parse_percent():
    assert parse_number("9%") == 9.0

def test_values_equal_numbers():
    assert values_equal("100.0", 100, tolerance=0.01)

def test_values_equal_strings():
    assert values_equal("abc", "abc")
