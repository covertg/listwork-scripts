"""End-to-end integration test for `parse_dartmouth_bu` using the committed
synthetic fixture. One test for the in-memory path, one for the write-to-disk
path."""

import shutil
from pathlib import Path

import pandas as pd

from parse_employer_bu import parse_dartmouth_bu

# Must match the columns in sample_list.xlsx.
ADDRESS_COLS = ("ADDRESS_LINE1", "ADDRESS_LINE2", "TOWN/CITY", "ST", "ZIP")
PROGRAM_COL = "PROGRAM"
NAME_COL = "NAME"

# The sample fixture filename embeds a date so get_list_identifier has something
# to parse. We copy the fixture into tmp_path with a dated filename.
FIXTURE_DATE = "2024.09.15"


def _stage_fixture(tmp_path: Path, sample_xlsx_path: Path) -> Path:
    dest = tmp_path / f"Sample BU {FIXTURE_DATE}.xlsx"
    shutil.copy(sample_xlsx_path, dest)
    return dest


def test_integration_full_pipeline(tmp_path, sample_xlsx_path, sample_mapping_path):
    infile = _stage_fixture(tmp_path, sample_xlsx_path)

    df = parse_dartmouth_bu(
        infile=infile,
        program_col=PROGRAM_COL,
        name_cols=[NAME_COL],
        program_mapping_file=sample_mapping_path,
        outfile=None,
        address_cols=ADDRESS_COLS,
        write=False,
    )

    expected_cols = {
        "Last",
        "First",
        "Middle",
        "Employer",
        "Degree",
        "Address Combined",
        f"BU List Employer {FIXTURE_DATE}",
    }
    assert expected_cols.issubset(df.columns)
    assert len(df) == 10  # same as sample_list.xlsx row count
    assert df[f"BU List Employer {FIXTURE_DATE}"].all()

    # Spot-check: row 0 is "Smith, John A.", program "Computer Science",
    # address "123 Fake Main St Apt 4, Hanover NH 03755".
    row0 = df.iloc[0]
    assert row0["Last"] == "Smith"
    assert row0["First"] == "John"
    assert row0["Middle"] == "A."
    assert row0["Employer"] == "FakeU CS"
    assert row0["Degree"] == "PhD"
    assert row0["Address Combined"] == "123 Fake Main St Apt 4, Hanover NH 03755"

    # Spot-check: row 4 is "O'Brien, Sean", no line2, "Mathematics" → FakeU MATH/PhD.
    row4 = df.iloc[4]
    assert row4["Last"] == "O'Brien"
    assert row4["First"] == "Sean"
    assert row4["Middle"] == ""
    assert row4["Employer"] == "FakeU MATH"
    assert row4["Address Combined"] == "654 Fake Birch Ln, West Lebanon NH 03784"

    # Program-suffix stripping should have collapsed "Biology PROGRAM" and
    # "Chemistry program" into their mapped values.
    assert df.iloc[1]["Employer"] == "FakeU BIO"
    assert df.iloc[2]["Employer"] == "FakeU CHEM"


def test_integration_write_to_disk(tmp_path, sample_xlsx_path, sample_mapping_path):
    infile = _stage_fixture(tmp_path, sample_xlsx_path)
    outfile = tmp_path / "output.csv"

    df = parse_dartmouth_bu(
        infile=infile,
        program_col=PROGRAM_COL,
        name_cols=[NAME_COL],
        program_mapping_file=sample_mapping_path,
        outfile=outfile,
        address_cols=ADDRESS_COLS,
        write=True,
    )

    assert outfile.exists()
    loaded = pd.read_csv(outfile)
    assert len(loaded) == len(df) == 10
    # Expected columns round-trip through the CSV.
    for col in ("Last", "First", "Employer", "Degree", "Address Combined"):
        assert col in loaded.columns
