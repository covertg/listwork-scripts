"""Tests for `parse_employer_bu.check_data_quality`.

These checks are advisory: they must never raise, and must never change the
dataframe. Each case below mirrors a defect actually seen in a Dartmouth
export.
"""

import numpy as np
import pandas as pd

from parse_employer_bu import check_data_quality


NAME_COLS = ["NAME"]
CITY, STATE, ZIP = "TOWN/CITY", "ST", "ZIP"


def _check(rows: list[dict], name_cols: list[str] = NAME_COLS) -> tuple[int, str]:
    df = pd.DataFrame(rows)
    before = df.copy()
    n = check_data_quality(df, name_cols=name_cols, cityc=CITY, statec=STATE, zipc=ZIP)
    pd.testing.assert_frame_equal(df, before)  # advisory only: no mutation
    return n, df


def _row(name="Smith, John", city="Hanover", st="NH", zipc="03755") -> dict:
    return {"NAME": name, CITY: city, STATE: st, ZIP: zipc}


def test_clean_data_flags_nothing(capsys):
    n, _ = _check(
        [_row(), _row(name="Doe, Jane", city="Norwich", st="VT", zipc="05055")]
    )
    assert n == 0
    assert "No data quality issues found" in capsys.readouterr().out


def test_zip_plus_four_is_not_flagged():
    n, _ = _check([_row(zipc="03784-1029")])
    assert n == 0


def test_duplicate_name_is_flagged(capsys):
    n, _ = _check([_row(), _row()])
    assert n == 2
    assert "share a name" in capsys.readouterr().out


def test_shifted_address_columns_flagged(capsys):
    # The real-world defect: city lands in the state column, state in the zip
    # column, and the zip in whatever comes next.
    n, _ = _check([_row(city=np.nan, st="Nashua", zipc="NH")])
    assert n == 1
    out = capsys.readouterr().out
    assert "shifted" in out


def test_international_address_flagged_but_distinctly(capsys):
    # No US state, non-US postcode. Legitimate, but worth eyeballing.
    n, _ = _check([_row(city="Shenzhen", st=np.nan, zipc="518000")])
    assert n == 1
    assert "zip" in capsys.readouterr().out


def test_missing_address_entirely_is_not_flagged():
    # A worker with no home address on file is a gap, not a corrupt row, and
    # load_list already reports null counts.
    n, _ = _check([_row(city=np.nan, st=np.nan, zipc=np.nan)])
    assert n == 0


def test_shifted_row_not_double_counted():
    # A shifted row also has a malformed state/zip; it should be reported once.
    n, _ = _check([_row(city=np.nan, st="Lebanon", zipc="NH")])
    assert n == 1


def test_lfm_columns_use_all_three_for_duplicate_detection():
    rows = [
        {
            "Last": "Smith",
            "First": "John",
            "Middle": "A.",
            CITY: "Hanover",
            STATE: "NH",
            ZIP: "03755",
        },
        {
            "Last": "Smith",
            "First": "Jane",
            "Middle": "B.",
            CITY: "Hanover",
            STATE: "NH",
            ZIP: "03755",
        },
    ]
    # Same last name, different people — must not be flagged as duplicates.
    n, _ = _check(rows, name_cols=["Last", "First", "Middle"])
    assert n == 0


def test_us_state_with_malformed_zip_flagged_separately(capsys):
    # A dropped leading zero. Seen in most historical lists; it is a typo in the
    # worker's record, not an international address, and should say so.
    n, _ = _check([_row(zipc="3755")])
    assert n == 1
    out = capsys.readouterr().out
    assert "US state but a zip" in out


def test_us_state_with_too_many_zip_digits_flagged(capsys):
    n, _ = _check([_row(zipc="037555")])
    assert n == 1
    assert "US state but a zip" in capsys.readouterr().out


def test_international_not_reported_as_zip_typo(capsys):
    n, _ = _check([_row(city="Auckland", st=np.nan, zipc="1010")])
    assert n == 1
    out = capsys.readouterr().out
    assert "international" in out
    assert "US state but a zip" not in out


def test_zip_typo_and_international_counted_once_each(capsys):
    n, _ = _check(
        [
            _row(name="Smith, John", zipc="3755"),
            _row(name="Doe, Jane", city="Auckland", st=np.nan, zipc="1010"),
        ]
    )
    assert n == 2
