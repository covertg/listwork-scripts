"""Tests for `parse_employer_bu.str_combine`."""

import math

from parse_employer_bu import str_combine


def test_str_combine_all_strings():
    assert str_combine("a", "b", "c") == "a b c"


def test_str_combine_some_empty():
    assert str_combine("a", "", "c") == "a c"


def test_str_combine_all_empty():
    assert str_combine("", "", "") == ""


def test_str_combine_with_nan():
    assert str_combine("a", math.nan, "c") == "a c"


def test_str_combine_with_none():
    assert str_combine("a", None, "c") == "a c"


def test_str_combine_trailing_commas():
    assert str_combine("a,", "b,") == "a b"


def test_str_combine_whitespace_stripped():
    assert str_combine("  a  ", "  b  ") == "a b"


def test_str_combine_no_args():
    assert str_combine() == ""
