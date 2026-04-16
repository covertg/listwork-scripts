"""Tests for `parse_employer_bu.make_address_combined`."""

import numpy as np
import pandas as pd

from parse_employer_bu import make_address_combined


ADDR_COLS = ("ADDRESS_LINE1", "ADDRESS_LINE2", "TOWN/CITY", "ST", "ZIP")


def _build_df(rows: list[dict]) -> pd.DataFrame:
    return pd.DataFrame(rows, columns=list(ADDR_COLS))


def _combine_one(row: dict) -> str:
    df = _build_df([row])
    out = make_address_combined(df, *ADDR_COLS)
    return out["Address Combined"].iloc[0]


def test_make_address_full():
    assert (
        _combine_one(
            {
                "ADDRESS_LINE1": "123 Fake Main St",
                "ADDRESS_LINE2": "Apt 4",
                "TOWN/CITY": "Hanover",
                "ST": "NH",
                "ZIP": "03755",
            }
        )
        == "123 Fake Main St Apt 4, Hanover NH 03755"
    )


def test_make_address_no_line2():
    assert (
        _combine_one(
            {
                "ADDRESS_LINE1": "123 Fake Main St",
                "ADDRESS_LINE2": "",
                "TOWN/CITY": "Hanover",
                "ST": "NH",
                "ZIP": "03755",
            }
        )
        == "123 Fake Main St, Hanover NH 03755"
    )


def test_make_address_no_line2_nan():
    assert (
        _combine_one(
            {
                "ADDRESS_LINE1": "123 Fake Main St",
                "ADDRESS_LINE2": np.nan,
                "TOWN/CITY": "Hanover",
                "ST": "NH",
                "ZIP": "03755",
            }
        )
        == "123 Fake Main St, Hanover NH 03755"
    )


def test_make_address_only_line1():
    # City/state/zip all empty strings → the post-comma branch is skipped.
    assert (
        _combine_one(
            {
                "ADDRESS_LINE1": "123 Fake Main St",
                "ADDRESS_LINE2": "",
                "TOWN/CITY": "",
                "ST": "",
                "ZIP": "",
            }
        )
        == "123 Fake Main St"
    )


def test_make_address_only_line1_with_nan_city_state_zip():
    # Edge case documenting current behavior: NaN is truthy in Python, so the
    # post-comma branch IS entered when any of town/st/zip are NaN, which
    # yields a trailing ", " after str_combine filters out the NaNs. In
    # practice this case is unreachable because load_list raises on empty
    # strings and leaves NaN only in ADDRESS_LINE2, but pinning the behavior
    # here prevents accidental drift.
    result = _combine_one(
        {
            "ADDRESS_LINE1": "123 Fake Main St",
            "ADDRESS_LINE2": "",
            "TOWN/CITY": np.nan,
            "ST": np.nan,
            "ZIP": np.nan,
        }
    )
    assert result == "123 Fake Main St, "


def test_make_address_trailing_commas_in_source():
    # Employer data sometimes has stray trailing commas on address parts;
    # str_combine strips them so the output shouldn't contain double commas.
    result = _combine_one(
        {
            "ADDRESS_LINE1": "123 Fake Main St,",
            "ADDRESS_LINE2": "Apt 4,",
            "TOWN/CITY": "Hanover",
            "ST": "NH",
            "ZIP": "03755",
        }
    )
    assert ",," not in result
    assert result == "123 Fake Main St Apt 4, Hanover NH 03755"


def test_make_address_preserves_zip_leading_zero():
    result = _combine_one(
        {
            "ADDRESS_LINE1": "123 Fake Main St",
            "ADDRESS_LINE2": "",
            "TOWN/CITY": "Hanover",
            "ST": "NH",
            "ZIP": "03755",
        }
    )
    assert "03755" in result
    assert "3755" not in result.replace("03755", "")  # no stray un-padded form


def test_make_address_preserves_row_count():
    df = _build_df(
        [
            {
                "ADDRESS_LINE1": "123 Fake Main St",
                "ADDRESS_LINE2": "Apt 4",
                "TOWN/CITY": "Hanover",
                "ST": "NH",
                "ZIP": "03755",
            },
            {
                "ADDRESS_LINE1": "456 Fake Oak Ave",
                "ADDRESS_LINE2": np.nan,
                "TOWN/CITY": "Norwich",
                "ST": "VT",
                "ZIP": "05055",
            },
            {
                "ADDRESS_LINE1": "789 Fake Pine Rd",
                "ADDRESS_LINE2": np.nan,
                "TOWN/CITY": "White River Junction",
                "ST": "VT",
                "ZIP": "05001",
            },
        ]
    )
    out = make_address_combined(df, *ADDR_COLS)
    assert len(out) == 3
    assert "Address Combined" in out.columns
