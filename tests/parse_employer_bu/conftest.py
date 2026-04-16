"""Shared fixtures for the parse_employer_bu test suite.

Also adds the repo root to `sys.path` so that tests can `import parse_employer_bu`
without needing a packaged install.
"""

from __future__ import annotations

import sys
import tomllib
from pathlib import Path

import pandas as pd
import pytest

REPO_ROOT = Path(__file__).resolve().parents[2]
FIXTURES_DIR = Path(__file__).parent / "fixtures"

# Make parse_employer_bu importable from tests.
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


# ---------- Path fixtures ----------


@pytest.fixture
def sample_xlsx_path() -> Path:
    return FIXTURES_DIR / "sample_list.xlsx"


@pytest.fixture
def sample_mapping_path() -> Path:
    return FIXTURES_DIR / "sample_program_mapping.toml"


@pytest.fixture
def real_fixtures_dir() -> Path:
    return FIXTURES_DIR / "real"


@pytest.fixture
def real_xlsx_files(real_fixtures_dir: Path) -> list[Path]:
    if not real_fixtures_dir.exists():
        return []
    return sorted(real_fixtures_dir.glob("*.xlsx"))


# ---------- Data fixtures ----------


@pytest.fixture
def sample_dataframe() -> pd.DataFrame:
    """Small in-memory DataFrame shaped like a BU list.

    Kept deliberately similar to `sample_list.xlsx` so unit tests that don't
    need file I/O can use this instead.
    """
    import numpy as np

    return pd.DataFrame(
        {
            "EMPLID": ["E0001", "E0002", "E0003"],
            "NAME": ["Smith, John A.", "Doe, Jane", "O'Brien, Sean"],
            "PROGRAM": ["Computer Science", "Biology PROGRAM", "Mathematics"],
            "ADDRESS_LINE1": [
                "123 Fake Main St",
                "456 Fake Oak Ave",
                "654 Fake Birch Ln",
            ],
            "ADDRESS_LINE2": ["Apt 4", np.nan, np.nan],
            "TOWN/CITY": ["Hanover", "Norwich", "West Lebanon"],
            "ST": ["NH", "VT", "NH"],
            "ZIP": ["03755", "05055", "03784"],
        }
    )


@pytest.fixture
def tmp_mapping_file(tmp_path: Path):
    """Factory fixture — write a mapping dict to a temp TOML file and return its path.

    Usage:
        def test_x(tmp_mapping_file):
            p = tmp_mapping_file({"Computer Science": ["CS", "PhD"]})
    """

    def _write(mapping: dict[str, list[str]]) -> Path:
        # Use tomllib-compatible hand-rolled output so we don't depend on tomli_w.
        # Values are always a 2-element list of strings, matching the production schema.
        lines = []
        for key, value in mapping.items():
            # Keys can contain spaces / punctuation — always quote them.
            # Escape any embedded double-quotes in keys (unlikely but harmless).
            safe_key = key.replace("\\", "\\\\").replace('"', '\\"')
            safe_vals = [
                v.replace("\\", "\\\\").replace('"', '\\"') for v in value
            ]
            vals_str = ", ".join(f'"{v}"' for v in safe_vals)
            lines.append(f'"{safe_key}" = [{vals_str}]')
        out = tmp_path / "mapping.toml"
        out.write_text("\n".join(lines) + "\n", encoding="utf-8")
        # Round-trip sanity check — if this raises, our writer has a bug.
        with open(out, "rb") as f:
            assert tomllib.load(f) == mapping
        return out

    return _write
