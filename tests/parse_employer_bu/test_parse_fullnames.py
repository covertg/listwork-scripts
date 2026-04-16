"""Tests for `parse_employer_bu.parse_fullnames`.

Some tests in this module are marked `xfail(strict=True)` because they codify
intended behavior for a known parsing bug in the current implementation: any
token ending in `.` is treated as a middle initial, which misclassifies
suffixes like "Jr." and "Sr.". When the bug is fixed, these tests should flip
to passing — at which point the strict-xfail will fail loudly and prompt
removal of the marker.
"""

import pandas as pd
import pytest

from parse_employer_bu import parse_fullnames


def _parse_single(fullname: str) -> tuple[str, str, str]:
    """Helper: run parse_fullnames on a one-row df and return (Last, First, Middle)."""
    df = pd.DataFrame({"NAME": [fullname]})
    out = parse_fullnames(df, "NAME")
    row = out.iloc[0]
    return row["Last"], row["First"], row["Middle"]


# ---------- Well-formed names ----------


def test_parse_fullnames_standard():
    assert _parse_single("Smith, John A.") == ("Smith", "John", "A.")


def test_parse_fullnames_no_middle():
    assert _parse_single("Smith, John") == ("Smith", "John", "")


def test_parse_fullnames_multiple_first_no_middle():
    assert _parse_single("Smith, John Paul") == ("Smith", "John Paul", "")


def test_parse_fullnames_multiple_first_with_middle():
    assert _parse_single("Smith, John Paul A.") == ("Smith", "John Paul", "A.")


def test_parse_fullnames_hyphenated_last():
    assert _parse_single("Smith-Jones, Alex") == ("Smith-Jones", "Alex", "")


def test_parse_fullnames_apostrophe_last():
    assert _parse_single("O'Brien, Sean") == ("O'Brien", "Sean", "")


def test_parse_fullnames_compound_last():
    assert _parse_single("García López, María") == ("García López", "María", "")


def test_parse_fullnames_suffix_iii():
    # "III" doesn't end with ".", so it ends up in the First column. This is
    # the correct (and intended) behavior; contrast with the xfail tests below
    # for "Jr." / "Sr." which currently get misclassified as middle initials.
    assert _parse_single("Smith, John III") == ("Smith", "John III", "")


# ---------- Suffix-handling tests (currently buggy) ----------


@pytest.mark.xfail(
    strict=True,
    reason=(
        "Known bug: tokens ending in '.' are treated as middle initial, so 'Jr.' "
        "is misclassified. Will pass once parse_fullnames is taught about suffixes."
    ),
)
def test_parse_fullnames_suffix_jr():
    assert _parse_single("Smith, John Jr.") == ("Smith", "John Jr.", "")


@pytest.mark.xfail(
    strict=True,
    reason=(
        "Known bug: tokens ending in '.' are treated as middle initial, so 'Sr.' "
        "is misclassified. Will pass once parse_fullnames is taught about suffixes."
    ),
)
def test_parse_fullnames_suffix_sr():
    assert _parse_single("Smith, John Sr.") == ("Smith", "John Sr.", "")


# ---------- Malformed input ----------


def test_parse_fullnames_multiple_commas_raises():
    with pytest.raises(ValueError):
        _parse_single("Smith, Jr., John")


def test_parse_fullnames_no_comma_raises():
    with pytest.raises(ValueError):
        _parse_single("Smith John")


def test_parse_fullnames_empty_string_raises():
    with pytest.raises(ValueError):
        _parse_single("")


def test_parse_fullnames_whitespace_only_raises():
    with pytest.raises(ValueError):
        _parse_single("   ")


def test_parse_fullnames_extra_whitespace_raises():
    # Documents current behavior: parse_fullnames is not whitespace-tolerant
    # against un-stripped input. The invertibility check reconstructs a
    # canonicalized "Last, First Middle" string and compares it to the input;
    # extra whitespace around or inside the name breaks equality and raises.
    # In production this is fine because load_list() strips whitespace first.
    with pytest.raises(ValueError):
        _parse_single("  Smith , John  ")


def test_parse_fullnames_invertibility_double_space_raises():
    # Another invertibility-check failure: internal double-space inside the
    # first name cannot round-trip through the reconstruction.
    with pytest.raises(ValueError):
        _parse_single("Smith, John  Paul")


def test_parse_fullnames_missing_column_raises():
    df = pd.DataFrame({"WRONG": ["Smith, John"]})
    with pytest.raises(ValueError):
        parse_fullnames(df, "NAME")


# ---------- Shape / batch properties ----------


def test_parse_fullnames_preserves_row_count():
    df = pd.DataFrame(
        {
            "NAME": [
                "Smith, John A.",
                "Doe, Jane",
                "O'Brien, Sean",
                "Smith-Jones, Alex T.",
                "García López, María",
            ]
        }
    )
    out = parse_fullnames(df, "NAME")
    assert len(out) == 5
    assert set(["Last", "First", "Middle"]).issubset(out.columns)
