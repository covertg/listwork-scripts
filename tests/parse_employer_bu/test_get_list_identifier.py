"""Tests for `parse_employer_bu.get_list_identifier`."""

from pathlib import Path

import pytest

from parse_employer_bu import get_list_identifier


def test_get_list_identifier_standard():
    assert (
        get_list_identifier(Path("BU List 2024.09.15.xlsx"))
        == "BU List Employer 2024.09.15"
    )


def test_get_list_identifier_no_date_raises():
    with pytest.raises(ValueError):
        get_list_identifier(Path("BU List.xlsx"))


def test_get_list_identifier_invalid_date_raises():
    # Feb 45 is not a valid date; extract_date returns None; get_list_identifier
    # should surface that as a ValueError.
    with pytest.raises(ValueError):
        get_list_identifier(Path("BU 2024.13.45.xlsx"))
