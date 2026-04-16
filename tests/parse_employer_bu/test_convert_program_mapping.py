"""Tests for `parse_employer_bu.convert_program_mapping`."""

import numpy as np
import pandas as pd
import pytest

from parse_employer_bu import convert_program_mapping


def _make_df(programs: list) -> pd.DataFrame:
    return pd.DataFrame({"PROGRAM": programs})


def test_convert_program_standard(tmp_mapping_file):
    mapping_path = tmp_mapping_file(
        {"Computer Science": ["FakeU CS", "PhD"], "Biology": ["FakeU BIO", "PhD"]}
    )
    df = _make_df(["Computer Science", "Biology", "Computer Science"])
    out = convert_program_mapping(df, "PROGRAM", mapping_path)
    assert list(out["Employer"]) == ["FakeU CS", "FakeU BIO", "FakeU CS"]
    assert list(out["Degree"]) == ["PhD", "PhD", "PhD"]


def test_convert_program_strips_program_suffix(tmp_mapping_file):
    mapping_path = tmp_mapping_file({"Computer Science": ["FakeU CS", "PhD"]})
    df = _make_df(["Computer Science PROGRAM"])
    out = convert_program_mapping(df, "PROGRAM", mapping_path)
    assert out["Employer"].iloc[0] == "FakeU CS"
    assert out["Degree"].iloc[0] == "PhD"


def test_convert_program_strips_program_suffix_case_insensitive(tmp_mapping_file):
    mapping_path = tmp_mapping_file({"Computer Science": ["FakeU CS", "PhD"]})
    df = _make_df(["Computer Science program"])
    out = convert_program_mapping(df, "PROGRAM", mapping_path)
    assert out["Employer"].iloc[0] == "FakeU CS"


def test_convert_program_unknown_raises(tmp_mapping_file):
    mapping_path = tmp_mapping_file({"Computer Science": ["FakeU CS", "PhD"]})
    df = _make_df(["Computer Science", "Astrology"])
    with pytest.raises(RuntimeError):
        convert_program_mapping(df, "PROGRAM", mapping_path)


def test_convert_program_null_raises(tmp_mapping_file):
    mapping_path = tmp_mapping_file({"Computer Science": ["FakeU CS", "PhD"]})
    df = _make_df(["Computer Science", np.nan])
    with pytest.raises(RuntimeError):
        convert_program_mapping(df, "PROGRAM", mapping_path)


def test_convert_program_missing_column_raises(tmp_mapping_file):
    mapping_path = tmp_mapping_file({"Computer Science": ["FakeU CS", "PhD"]})
    df = pd.DataFrame({"WRONG": ["Computer Science"]})
    with pytest.raises(ValueError):
        convert_program_mapping(df, "PROGRAM", mapping_path)


def test_convert_program_missing_mapping_file_raises(tmp_path):
    df = _make_df(["Computer Science"])
    with pytest.raises(FileNotFoundError):
        convert_program_mapping(df, "PROGRAM", tmp_path / "does_not_exist.toml")


def test_convert_program_preserves_row_count(tmp_mapping_file):
    mapping_path = tmp_mapping_file(
        {"Computer Science": ["FakeU CS", "PhD"], "Biology": ["FakeU BIO", "PhD"]}
    )
    df = _make_df(
        ["Computer Science", "Biology", "Computer Science", "Biology", "Biology"]
    )
    out = convert_program_mapping(df, "PROGRAM", mapping_path)
    assert len(out) == 5
