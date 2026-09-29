"""Generate `sample_list.xlsx`, the small synthetic employer BU list used by the
parse_employer_bu test suite.

Run this script to (re)create the fixture:

    python tests/parse_employer_bu/fixtures/generate_sample.py

The output is deterministic. All names and addresses are obviously fake.

Column naming notes (real employer lists vary — document here when they change):
- `EMPLID` is a placeholder for the employee ID column (real column name varies).
- `NAME` is the placeholder for the fullname column (real column name is
  usually something like "FULL_NAME (LFM)").
- `PROGRAM` is the placeholder for the program / field-of-study column (real
  column name is usually "Program/Field of Study", sometimes uppercased).
- The address columns match the real employer column names as of WI26:
  ADDRESS_LINE1, ADDRESS_LINE2, TOWN/CITY, ST, ZIP.
"""

from pathlib import Path

import pandas as pd


OUT_PATH = Path(__file__).parent / "sample_list.xlsx"


# Each row exercises a specific branch of the parsing logic. Columns intentionally
# match the employer's real column naming for address fields. See module docstring
# for notes on the placeholder columns (EMPLID, NAME, PROGRAM).
#
# Coverage:
#   1: standard name w/ middle initial + standard program + full address (line2 present)
#   2: name with no middle + program with " PROGRAM" suffix + no line2
#   3: name with multiple first names + lowercase "program" suffix
#   4: compound last name with diacritics (no middle, no line2)
#   5: apostrophe in last name
#   6: hyphenated last name + middle initial + standard program
#   7: name with multiple first names AND middle initial
#   8: single first name + standard program
#   9: zip code with leading zero (already covered above; this row double-checks)
#  10: another plain row
ROWS = [
    {
        "EMPLID": "E0001",
        "NAME": "Smith, John A.",
        "PROGRAM": "Computer Science",
        "ADDRESS_LINE1": "123 Fake Main St",
        "ADDRESS_LINE2": "Apt 4",
        "TOWN/CITY": "Hanover",
        "ST": "NH",
        "ZIP": "03755",
    },
    {
        "EMPLID": "E0002",
        "NAME": "Doe, Jane",
        "PROGRAM": "Biology PROGRAM",
        "ADDRESS_LINE1": "456 Fake Oak Ave",
        "ADDRESS_LINE2": None,  # -> NaN after read_excel; tests line2-missing branch
        "TOWN/CITY": "Norwich",
        "ST": "VT",
        "ZIP": "05055",
    },
    {
        "EMPLID": "E0003",
        "NAME": "Brown, Maria Elena",
        "PROGRAM": "Chemistry program",  # lowercase suffix -> case-insensitive strip
        "ADDRESS_LINE1": "789 Fake Pine Rd",
        "ADDRESS_LINE2": None,
        "TOWN/CITY": "White River Junction",
        "ST": "VT",
        "ZIP": "05001",
    },
    {
        "EMPLID": "E0004",
        "NAME": "García López, José",
        "PROGRAM": "Physics",
        "ADDRESS_LINE1": "321 Fake Elm St",
        "ADDRESS_LINE2": None,
        "TOWN/CITY": "Lebanon",
        "ST": "NH",
        "ZIP": "03766",
    },
    {
        "EMPLID": "E0005",
        "NAME": "O'Brien, Sean",
        "PROGRAM": "Mathematics",
        "ADDRESS_LINE1": "654 Fake Birch Ln",
        "ADDRESS_LINE2": None,
        "TOWN/CITY": "West Lebanon",
        "ST": "NH",
        "ZIP": "03784",
    },
    {
        "EMPLID": "E0006",
        "NAME": "Smith-Jones, Alex T.",
        "PROGRAM": "Engineering PROGRAM",
        "ADDRESS_LINE1": "987 Fake Maple Dr",
        "ADDRESS_LINE2": "Unit 5",
        "TOWN/CITY": "Enfield",
        "ST": "NH",
        "ZIP": "03748",
    },
    {
        "EMPLID": "E0007",
        "NAME": "Johnson, Robert Paul J.",
        "PROGRAM": "Computer Science",
        "ADDRESS_LINE1": "147 Fake Cedar Way",
        "ADDRESS_LINE2": None,
        "TOWN/CITY": "Etna",
        "ST": "NH",
        "ZIP": "03750",
    },
    {
        "EMPLID": "E0008",
        "NAME": "Williams, Patricia",
        "PROGRAM": "Music",
        "ADDRESS_LINE1": "258 Fake Spruce Ct",
        "ADDRESS_LINE2": None,
        "TOWN/CITY": "Hartford",
        "ST": "VT",
        "ZIP": "05047",
    },
    {
        "EMPLID": "E0009",
        "NAME": "Davis, Michael",
        "PROGRAM": "History",
        "ADDRESS_LINE1": "369 Fake Walnut Blvd",
        "ADDRESS_LINE2": "Apt 12",
        "TOWN/CITY": "Hanover",
        "ST": "NH",
        "ZIP": "03755",
    },
    {
        "EMPLID": "E0010",
        "NAME": "Zhang, Wei",
        "PROGRAM": "Physics",
        "ADDRESS_LINE1": "741 Fake Cherry St",
        "ADDRESS_LINE2": None,
        "TOWN/CITY": "Lyme",
        "ST": "NH",
        "ZIP": "03768",
    },
]


def build_dataframe() -> pd.DataFrame:
    df = pd.DataFrame(ROWS)
    # Force all columns to object/string dtype so that the zip codes are written
    # as text (not numbers). This is what the employer's files also do in practice.
    return df.astype(object)


def write_sample(path: Path = OUT_PATH) -> Path:
    df = build_dataframe()
    path.parent.mkdir(parents=True, exist_ok=True)
    # Use openpyxl explicitly so we can ensure leading-zero zip codes survive.
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="BU")
        # Force the ZIP column to Text format so openpyxl preserves leading zeros.
        ws = writer.sheets["BU"]
        zip_col_idx = df.columns.get_loc("ZIP") + 1  # 1-based
        for row in ws.iter_rows(
            min_row=2, max_row=ws.max_row, min_col=zip_col_idx, max_col=zip_col_idx
        ):
            for cell in row:
                cell.number_format = "@"  # Text format
                if cell.value is not None:
                    cell.value = str(cell.value)
    return path


if __name__ == "__main__":
    out = write_sample()
    print(f"Wrote {out}")
