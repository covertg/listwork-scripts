"""Tests for `parse_employer_bu.check_columns`.

This is the fail-fast guard that catches mistyped `--program_col` /
`--fullname_col` / `--address_cols` arguments before any parsing happens.
"""

import pandas as pd
import pytest

from parse_employer_bu import check_columns


DF = pd.DataFrame({"NAME": ["Smith, John"], "PROGRAM": ["Biology"]})


def test_check_columns_all_present():
    # Should return quietly.
    check_columns(DF, {"--program_col": "PROGRAM", "--fullname_col": "NAME"})


def test_check_columns_missing_raises():
    with pytest.raises(ValueError):
        check_columns(DF, {"--program_col": "Program1 Code"})


def test_check_columns_error_names_the_missing_column(capsys):
    with pytest.raises(ValueError) as excinfo:
        check_columns(DF, {"--program_col": "Program1 Code"})
    assert "Program1 Code" in str(excinfo.value)
    # The real columns are printed so the officer can see what to use instead.
    out = capsys.readouterr().out
    assert "'NAME'" in out
    assert "'PROGRAM'" in out


def test_check_columns_reports_every_missing_column(capsys):
    with pytest.raises(ValueError):
        check_columns(
            DF,
            {
                "--program_col": "PROGRAM",  # present
                "--fullname_col": "LFM Name Formatted",  # missing
                "--address_cols (city)": "LO City",  # missing
            },
        )
    out = capsys.readouterr().out
    assert "LFM Name Formatted" in out
    assert "LO City" in out
