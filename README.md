# listwork-scripts

## `parse_employer_bu.py`

Process a BU list from Dartmouth for the purposes of uploading to Broadstripes (and otherwise using). The main complicated part of this process is mapping the hard-to-interpret program/field of study codes from Dartmouth to program and degree type fields that we can more easily use. This process also does a few other things that are nice to automate, like combining multiple address columns into one column, and separating a unified name column (given as “Last First Middle”) into separate name columns.

Requires: `pandas`, `openpyxl`, `python>=3.11`

**If modifying this script, check and potentially add tests. See `tests/parse_employer_bu/README.md`.**

Usage info:
```bash
python parse_employer_bu.py --help
```
```bash
# output
usage: parse_employer_bu.py [-h] -i INFILE [-o OUTFILE] --program_col PROGRAM_COL [--fullname_col FULLNAME_COL] [--lfm_cols LFM_COLS LFM_COLS LFM_COLS]
                            [--program_mapping_file PROGRAM_MAPPING_FILE] [--address_cols LINE1 LINE2 CITY STATE ZIP]

Parse a Dartmouth BU list into a CSV file, cleaned and formatted to use with Broadstripes.

options:
  -h, --help            show this help message and exit
  -i, --infile INFILE   Path to the input file (.xlsx from employer). The filename must include the date that we received it in the format YYYY.MM.DD
  -o, --outfile OUTFILE
                        Path to the output file (.csv). Optional. By default, the output file will be in the data/ directory and have an informative name
                        with a timestamp.
  --program_col PROGRAM_COL
                        Name of the column containing the program/field of study. This column name seems to vary pretty frequently by term, so you will
                        need to identify it by peeking at the input file.
  --fullname_col FULLNAME_COL
                        Name of the column containing the full name. Only fullname_Col or lfm_cols can be specified. All of recent BU lists from
                        Dartmouth have used a unified column for the full name, so this is probably the argument you want for a new BU list.
  --lfm_cols LFM_COLS LFM_COLS LFM_COLS
                        (Legacy) Name of the columns containing the last name, first name, and middle initial, separated by spaces. Only lfm_cols or
                        fullname_col can be specified. Dartmouth's original BU lists separated names into columns, but more recent lists have all used a
                        unified column, so you probably don't want this argument for a new BU list.
  --program_mapping_file PROGRAM_MAPPING_FILE
                        Path to the program mapping file (.toml that we maintain). Defaults to ./program_mapping.toml
  --address_cols LINE1 LINE2 CITY STATE ZIP
                        Names of the address columns: line1, line2, city, state, and zip (in that order).
```

Example usage:
```bash
python parse_employer_bu.py -i './data/2026.04.16 TO GOLD Membership 26S.xlsx' --program_col "Program1 Code" --fullname_col "LFM Name Formatted" --address_cols "LO Street1" "LO Street2" "LO City" "LO State" "LO Zip"
```

```bash
# output
Loaded file './data/2026.04.16 TO GOLD Membership 26S.xlsx'.
n. rows:         868
Columns:         ['LFM Name Formatted', 'Chosen-Pref-Legal First Name', 'DND Email', 'LO Street1', 'LO Street2', 'LO City', 'LO State', 'LO Zip', 'MP Phone For Address', 'Program1 Code']
Column(s) with null values:
            % null  total null
LO Street2   38.13         331
LO State      0.12           1

Parsed employer list date as '2026.04.16'.

Parsing program/field of study data...
Parsed 25 different programs/departments and 11 different degree types.

Parsed full name -> First Last Middle.

Combined address column created.

Finished parsing this employer BU list. Please check the output for errors before using it.
Wrote to file 'data/BU List Employer 2026.04.16 made 2026.04.29_17.05.53.csv'
```

## `check_skipped_imports.py`

Once we've processed a BU list, we first update only the workers that already exist in Broadstripes. Broadstripes gives us a list of "skips" which should represent the workers that are new to us. However, sometimes a worker just changes their name or email. To avoid these false duplicates, we compare the "skips" to all of the entries in our Broadstripes database and search for names that approximately match.

This task will never be done perfectly, and that's fine. But if we wanted to improve it some day, we could also try matching entries based on other information, e.g. phone number and address seem promising.

Requires: `pandas`

This script has a light set of tests. If modifying it, check and potentially add tests. See `tests/check_skipped_imports/`.

Usage info:
```bash
python check_skipped_imports.py
```

```bash
# output
# TODO
```

Example usage (outdated):
```bash
# TODO
```