"""Tests for check_skipped_imports.py, using small synthetic CSVs.

Run from the repo root with `pytest tests/`.
"""

import sys
from pathlib import Path

import pandas as pd
import pytest

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from check_skipped_imports import (  # noqa: E402
    check_skipped_imports,
    find_potential_matches,
    format_name,
    main,
    normalize_phone,
)


# ---------- Unit tests ----------


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("(603) 555-1234", "6035551234"),
        ("603-555-1234", "6035551234"),
        ("6035551234", "6035551234"),
        ("+16035551234", "6035551234"),
        ("603.555.1234", "6035551234"),
        ("+34673800685", "34673800685"),  # international: keep as-is
        ("12", ""),  # too short to trust
        (None, ""),
        (float("nan"), ""),
    ],
)
def test_normalize_phone(raw, expected):
    assert normalize_phone(raw) == expected


def test_format_name():
    assert format_name("Smith", "Jane", "Q.") == "Smith, Jane Q."
    assert format_name("Smith", "Jane", float("nan")) == "Smith, Jane"
    assert format_name(" Smith ", "Jane", None) == "Smith, Jane"
    assert format_name(None, None, None) == ""


def test_exact_name_match_is_case_insensitive():
    groups = find_potential_matches(["Smith, Jane"], ["smith, jane"], [""], [""])
    assert len(groups) == 1
    assert groups[0]["strong"]
    assert groups[0]["matches"][0]["reasons"] == ["same name"]


def test_similar_name_respects_threshold():
    skip, bs = ["Sarathchandra Andarage, Sanduni"], ["Sarathchandra, Sanduni"]
    assert find_potential_matches(skip, bs, [""], [""], threshold=0.75)
    assert not find_potential_matches(skip, bs, [""], [""], threshold=0.95)


def test_same_phone_catches_a_changed_name():
    groups = find_potential_matches(
        ["Petrov, Beverly"], ["Oldname, Beverly"], ["6035551234"], ["6035551234"],
        threshold=0.95,
    )
    assert len(groups) == 1
    assert groups[0]["strong"]
    assert groups[0]["matches"][0]["reasons"] == ["same phone"]


def test_same_phone_and_similar_name_are_combined():
    groups = find_potential_matches(
        ["Smith, Jane"], ["Smyth, Jane"], ["6035551234"], ["6035551234"]
    )
    assert groups[0]["matches"][0]["reasons"] == ["same phone", "similar name"]


def test_missing_phones_never_match_each_other():
    groups = find_potential_matches(["Aaa, Bbb"], ["Xxx, Yyy"], [""], [""])
    assert groups == []


def test_strong_groups_sort_before_weaker_ones():
    skips = ["Smith, Jon", "Doe, Jane"]
    bs = ["Smith, John", "Doe, Jane"]
    groups = find_potential_matches(skips, bs, ["", ""], ["", ""])
    assert [g["skip_idx"] for g in groups] == [1, 0]
    assert [g["strong"] for g in groups] == [True, False]


# ---------- End-to-end ----------


@pytest.fixture
def csvs(tmp_path: Path) -> tuple[Path, Path]:
    skips = pd.DataFrame(
        {
            "Last": ["Petrov", "Brandnew", "Smith"],
            "First": ["Beverly", "Person", "Jon"],
            "Middle": [None, None, None],
            "MP Phone For Address": ["(603) 555-1234", "(603) 555-0000", None],
            "DND Email": ["b.petrov@x.edu", "p.brandnew@x.edu", "j.smith@x.edu"],
            "Employer": ["ENGS", "COSC", "TUCK"],
        }
    )
    bs = pd.DataFrame(
        {
            "Broadstripes ID": ["AAA-111", "BBB-222", "CCC-333"],
            "Last Name": ["Allen", "Smith", "Unrelated"],
            "First Name": ["Beverly", "John", "Someone"],
            "Middle Name": [None, None, None],
            "Primary Phone": ["6035551234", None, "6035559999"],
            "Primary Email": ["b.allen@x.edu", "john.smith@x.edu", "u@x.edu"],
            "Employer": ["ENGS", "TUCK", "MATH"],
        }
    )
    skips_path = tmp_path / "data-import-SKIPS-test.csv"
    bs_path = tmp_path / "everyone.csv"
    skips.to_csv(skips_path, index=False)
    bs.to_csv(bs_path, index=False)
    return skips_path, bs_path


def test_end_to_end_writes_csv_next_to_skips(csvs, capsys):
    skips_path, bs_path = csvs
    tables = check_skipped_imports(skips_path, bs_path)

    out = capsys.readouterr().out
    assert "1 skips have no possible match" in out
    assert "1 strong" in out and "1 weaker" in out

    assert len(tables) == 2
    first = tables[0]
    assert first.iloc[0]["Name"] == "Petrov, Beverly"
    assert first.iloc[1]["Broadstripes ID"] == "AAA-111"
    assert first.iloc[1][""] == "Match: same phone"

    written = skips_path.parent / f"{skips_path.stem} - possible duplicates.csv"
    assert written.exists()
    # 2 skip rows + 2 match rows + 1 blank separator
    assert len(pd.read_csv(written, dtype=str)) == 5


def test_missing_column_lists_actual_columns(csvs, capsys):
    skips_path, bs_path = csvs
    with pytest.raises(ValueError, match="Phone"):
        check_skipped_imports(
            skips_path,
            bs_path,
            bs_info_cols=("Phone", "Primary Email", "Employer"),
            write=False,
        )
    out = capsys.readouterr().out
    assert "--broadstripes_info_cols (phone): 'Phone'" in out
    assert "'Primary Phone'" in out


def test_cli_rejects_bad_threshold(csvs):
    skips_path, bs_path = csvs
    with pytest.raises(SystemExit):
        main(["-s", str(skips_path), "-b", str(bs_path), "--threshold", "1.5"])


def test_cli_outfile(csvs, tmp_path):
    skips_path, bs_path = csvs
    outfile = tmp_path / "sub" / "out.csv"
    main(["-s", str(skips_path), "-b", str(bs_path), "-o", str(outfile)])
    assert outfile.exists()
