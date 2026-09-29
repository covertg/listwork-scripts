"""Tests for the command-line front end (`build_parser` / `main`).

These cover the argument validation that a new officer is most likely to trip
over, and the `--show_columns` shortcut for inspecting an unfamiliar file.
"""

import shutil
from pathlib import Path

import pandas as pd
import pytest

from parse_employer_bu import main


FIXTURE_DATE = "2024.09.15"
ARGS_COMMON = [
    "--program_col",
    "PROGRAM",
    "--fullname_col",
    "NAME",
    "--address_cols",
    "ADDRESS_LINE1",
    "ADDRESS_LINE2",
    "TOWN/CITY",
    "ST",
    "ZIP",
]


@pytest.fixture
def staged_xlsx(tmp_path: Path, sample_xlsx_path: Path) -> Path:
    dest = tmp_path / f"Sample BU {FIXTURE_DATE}.xlsx"
    shutil.copy(sample_xlsx_path, dest)
    return dest


def test_show_columns_prints_columns_and_exits(staged_xlsx, capsys):
    main(["-i", str(staged_xlsx), "--show_columns"])
    out = capsys.readouterr().out
    assert "PROGRAM" in out
    assert "ADDRESS_LINE1" in out
    # Nothing was written.
    assert not list(staged_xlsx.parent.glob("*.csv"))


def test_show_columns_does_not_require_program_col(staged_xlsx):
    # The whole point: you can peek before you know any column names.
    main(["-i", str(staged_xlsx), "--show_columns"])


def test_missing_program_col_errors(staged_xlsx):
    with pytest.raises(SystemExit):
        main(["-i", str(staged_xlsx), "--fullname_col", "NAME"])


def test_both_name_args_errors(staged_xlsx):
    with pytest.raises(SystemExit):
        main(
            [
                "-i",
                str(staged_xlsx),
                "--program_col",
                "PROGRAM",
                "--fullname_col",
                "NAME",
                "--lfm_cols",
                "LAST",
                "FIRST",
                "MIDDLE",
            ]
        )


def test_neither_name_arg_errors(staged_xlsx):
    with pytest.raises(SystemExit):
        main(["-i", str(staged_xlsx), "--program_col", "PROGRAM"])


def test_default_outfile_lands_next_to_infile(
    staged_xlsx, sample_mapping_path, monkeypatch, tmp_path
):
    # Regression guard: the default output path used to be a hard-coded "data/"
    # relative to the working directory, which broke when the script was run
    # from anywhere but the repo root.
    elsewhere = tmp_path / "elsewhere"
    elsewhere.mkdir()
    monkeypatch.chdir(elsewhere)

    main(
        [
            "-i",
            str(staged_xlsx),
            "--program_mapping_file",
            str(sample_mapping_path),
            *ARGS_COMMON,
        ]
    )

    written = list(staged_xlsx.parent.glob("BU List Employer *.csv"))
    assert len(written) == 1
    assert f"BU List Employer {FIXTURE_DATE}" in written[0].name
    assert len(pd.read_csv(written[0])) == 10


def test_mistyped_column_fails_before_writing_anything(
    staged_xlsx, sample_mapping_path
):
    with pytest.raises(ValueError):
        main(
            [
                "-i",
                str(staged_xlsx),
                "--program_mapping_file",
                str(sample_mapping_path),
                "--program_col",
                "Program1 Code",  # not in the fixture
                "--fullname_col",
                "NAME",
            ]
        )
    assert not list(staged_xlsx.parent.glob("*.csv"))
