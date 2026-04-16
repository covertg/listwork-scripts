"""Tests that run against real employer BU lists dropped into `fixtures/real/`.

If the directory is empty (or missing a `program_mapping.toml`), these tests
are skipped. See `fixtures/real/README.md` for how to use this.
"""

from __future__ import annotations

import tomllib
from pathlib import Path

import pandas as pd
import pytest

from parse_employer_bu import (
    DEFAULT_ADDRESS_COLS,
    extract_date,
    parse_dartmouth_bu,
)

REAL_DIR = Path(__file__).parent / "fixtures" / "real"
REAL_MAPPING = REAL_DIR / "program_mapping.toml"
REAL_FILES = sorted(REAL_DIR.glob("*.xlsx")) if REAL_DIR.exists() else []


# Sensible defaults matching the employer's most recent BU-list format.
DEFAULT_PARAMS = {
    "program_col": "Program/Field of Study",
    "fullname_col": "FULL_NAME (LFM)",
    "address_cols": list(DEFAULT_ADDRESS_COLS),
}


def _load_params(xlsx: Path) -> dict:
    """Merge DEFAULT_PARAMS with any per-file overrides in <xlsx-stem>.params.toml."""
    override = xlsx.with_suffix(".params.toml")
    params = dict(DEFAULT_PARAMS)
    if override.exists():
        with open(override, "rb") as f:
            params.update(tomllib.load(f))
    return params


def _run(xlsx: Path) -> pd.DataFrame:
    params = _load_params(xlsx)
    if "lfm_cols" in params:
        name_cols = list(params["lfm_cols"])
    elif params.get("fullname_col"):
        name_cols = [params["fullname_col"]]
    else:
        pytest.skip(f"No name columns configured for {xlsx.name}")
    return parse_dartmouth_bu(
        infile=xlsx,
        program_col=params["program_col"],
        name_cols=name_cols,
        program_mapping_file=REAL_MAPPING,
        outfile=None,
        address_cols=tuple(params["address_cols"]),
        write=False,
    )


_skip_no_files = pytest.mark.skipif(
    not REAL_FILES,
    reason=(
        "No real fixtures in tests/parse_employer_bu/fixtures/real/. "
        "See that dir's README.md for how to add them."
    ),
)
_skip_no_mapping = pytest.mark.skipif(
    not REAL_MAPPING.exists(),
    reason=(
        "No program_mapping.toml in tests/parse_employer_bu/fixtures/real/. "
        "Symlink or copy the production one."
    ),
)


@_skip_no_files
@_skip_no_mapping
@pytest.mark.parametrize(
    "real_file", REAL_FILES, ids=[f.name for f in REAL_FILES]
)
def test_real_file_parses_without_error(real_file):
    # Sanity: filename must contain a parseable date or parse_dartmouth_bu will
    # raise inside get_list_identifier. Surface a clearer error if not.
    assert extract_date(real_file.name), (
        f"Filename {real_file.name!r} is missing a YYYY.MM.DD date — the script "
        "requires this. Rename the file before testing."
    )
    _ = _run(real_file)


@_skip_no_files
@_skip_no_mapping
@pytest.mark.parametrize(
    "real_file", REAL_FILES, ids=[f.name for f in REAL_FILES]
)
def test_real_file_expected_columns(real_file):
    df = _run(real_file)
    expected = {"Last", "First", "Middle", "Employer", "Degree", "Address Combined"}
    missing = expected - set(df.columns)
    assert not missing, f"Missing columns: {missing}"


@_skip_no_files
@_skip_no_mapping
@pytest.mark.parametrize(
    "real_file", REAL_FILES, ids=[f.name for f in REAL_FILES]
)
def test_real_file_no_data_loss(real_file):
    # Row count preserved from load_list through to the returned frame.
    loaded = pd.read_excel(real_file, dtype=str)
    out = _run(real_file)
    assert len(out) == len(loaded)
