# listwork-scripts

## `parse_employer_bu.py`

Requires: `pandas`, `openpyxl`, `python>=3.11`

If modifying this script, check and potentially add tests. See `tests/parse_employer_bu/README.md`.

```bash
python parse_employer_bu.py --help
```
```bash
# output
usage: parse_employer_bu.py [-h] -i INFILE --program_col PROGRAM_COL [--fullname_col FULLNAME_COL] [--lfm_cols LFM_COLS LFM_COLS LFM_COLS]
                            [--program_mapping_file PROGRAM_MAPPING_FILE] [-o OUTFILE] [--address_cols LINE1 LINE2 CITY STATE ZIP]

Parse a Dartmouth BU list into a CSV file, cleaned and formatted to use with Broadstripes.

options:
  -h, --help            show this help message and exit
  -i, --infile INFILE   Path to the input file (.xlsx from employer). The filename must include the date that we received it in the format YYYY.MM.DD
  --program_col PROGRAM_COL
                        Name of the column containing the program/field of study. This column name seems to vary pretty frequently by term, so you will need to
                        identify it by peeking at the input file.
  --fullname_col FULLNAME_COL
                        Name of the column containing the full name. Only fullname_Col or lfm_cols can be specified. All of recent BU lists from Dartmouth have
                        used a unified column for the full name, so this is probably the argument you want for a new BU list.
  --lfm_cols LFM_COLS LFM_COLS LFM_COLS
                        (Legacy) Name of the columns containing the last name, first name, and middle initial, separated by spaces. Only lfm_cols or
                        fullname_col can be specified. Dartmouth's original BU lists separated names into columns, but more recent lists have all used a
                        unified column, so you probably don't want this argument for a new BU list.
  --program_mapping_file PROGRAM_MAPPING_FILE
                        Path to the program mapping file (.toml that we develop).
  -o, --outfile OUTFILE
                        Path to the output file (.csv). Optional. By default, the output file will be in the data/ directory and have an informative name with
                        a timestamp.
  --address_cols LINE1 LINE2 CITY STATE ZIP
                        Names of the address columns: line1, line2, city, state, and zip (in that order).
```

Example usage:
```bash
python parse_employer_bu.py -i "data/2025.09.26 TO GOLD Membership 25F.xlsx" --program_col "Program/Field of Study" --fullname_col "FULL_NAME (LFM)"
```

```bash
# output
Loaded file 'data/2025.09.26 TO GOLD Membership 25F.xlsx'.
n. rows:         835
Columns:         ['FULL_NAME (LFM)', 'CHOSEN FIRST NAME', 'EMAIL_ADDRESS', 'ADDRESS_LINE1', 'ADDRESS_LINE2', 'TOWN/CITY', 'ST', 'ZIP', 'PHONE', 'Program/Field of Study']
Column(s) with null values:
                   % null  total null
CHOSEN FIRST NAME    0.24           2
ADDRESS_LINE1        1.92          16
ADDRESS_LINE2       55.33         462
TOWN/CITY            1.80          15
ST                   1.80          15
ZIP                  1.80          15
PHONE              100.00         835

Parsed employer list date as '2025.09.26'.

Parsing program/field of study data...
Parsed 23 different programs/departments and 11 different degree types.

Parsed full name -> First Last Middle.

Combined address column created.

Finished parsing this employer BU list. Please check the output for errors before using it.
Wrote to file 'data/BU List Employer 2025.09.26 made 2026.04.16_15.08.12.csv'
```

## `check_skipped_imports.py`

Requires: `pandas`

Note: this script is currently good enough for us, but it could miss some cases. E.g. if someone entirely changes their name then it may not detect that. To be more comprehensive we could try to cross-reference by department and year.

Example usage:

```bash
python check_skipped_imports.py --all_broadstripes 'data/Basic contact info.csv' --skipped_entries data/data-import-SKIPS-d88d986d-19c5-4217-9a16-f78cf79540b1.csv
```

```bash
Loaded all Broadstripes entries from 'data/Basic contact info.csv'
n. rows:         2179
Loaded skipped additions from 'data/data-import-SKIPS-d88d986d-19c5-4217-9a16-f78cf79540b1.csv'
n. rows:         20

Found 2 potential matches
  name match_type  similarity       existing_matches
0 redacted fuzzy   0.918919         [redacted]
1 redacted fuzzy   0.826087         [redacted]
