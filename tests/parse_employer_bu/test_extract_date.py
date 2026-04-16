"""Tests for `parse_employer_bu.extract_date`."""

from parse_employer_bu import extract_date


def test_extract_date_standard():
    assert extract_date("BU List 2024.09.15.xlsx") == "2024.09.15"


def test_extract_date_at_start():
    assert extract_date("2024.09.15 BU List.xlsx") == "2024.09.15"


def test_extract_date_at_end():
    assert extract_date("BU List.2024.09.15") == "2024.09.15"


def test_extract_date_no_date():
    assert extract_date("BU List.xlsx") is None


def test_extract_date_invalid_month():
    assert extract_date("BU 2024.13.15.xlsx") is None


def test_extract_date_invalid_day():
    assert extract_date("BU 2024.09.32.xlsx") is None


def test_extract_date_leap_year_valid():
    assert extract_date("BU 2024.02.29.xlsx") == "2024.02.29"


def test_extract_date_leap_year_invalid():
    assert extract_date("BU 2023.02.29.xlsx") is None


def test_extract_date_wrong_separator():
    # Current regex requires literal dots as separators.
    assert extract_date("BU 2024-09-15.xlsx") is None


def test_extract_date_short_year():
    assert extract_date("BU 24.09.15.xlsx") is None


def test_extract_date_short_month():
    assert extract_date("BU 2024.9.15.xlsx") is None


def test_extract_date_multiple_dates_returns_first():
    # Documented intended behavior: re.search returns the first match.
    assert (
        extract_date("BU 2024.09.15 updated 2024.10.01.xlsx") == "2024.09.15"
    )


def test_extract_date_empty_string():
    assert extract_date("") is None
