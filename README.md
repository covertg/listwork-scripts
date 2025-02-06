# listwork-scripts

## `parse_employer_bu.py`

Requires: `pandas`, `openpyxl`

```bash
python parse_employer_bu.py --help
```
```bash
# output
usage: parse_employer_bu.py [-h] --infile INFILE --program_col PROGRAM_COL [--fullname_col FULLNAME_COL] [--lfm_cols LFM_COLS LFM_COLS LFM_COLS] [--program_mapping_file PROGRAM_MAPPING_FILE] [--outfile OUTFILE]

Parse a Dartmouth BU list into a CSV file, cleaned and formatted to use with Broadstripes.

options:
  -h, --help            show this help message and exit
  --infile INFILE       Path to the input file (.xlsx from employer)
  --program_col PROGRAM_COL
                        Name of the column containing the program/field of study
  --fullname_col FULLNAME_COL
                        Name of the column containing the full name. Only fullname_Col or lfm_cols can be specified.
  --lfm_cols LFM_COLS LFM_COLS LFM_COLS
                        Name of the columns containing the last name, first name, and middle initial, separated by spaces. Only lfm_cols or fullname_col can be specified.
  --program_mapping_file PROGRAM_MAPPING_FILE
                        Path to the program mapping file (.toml from us)
  --outfile OUTFILE     Path to the output file (.csv). Optional. By default, the output file will be in the data/ directory and have an informative name with a timestamp.

```

Example usage:

```bash
python parse_employer_bu.py --infile="data/GOLD Membership (Winter 2025) 2025.02.03.xlsx" --program_col="Program/Field of Study" --fullname_col="FULL_NAME (LFM)"
```

```bash
# output
Loaded file 'data/GOLD Membership (Winter 2025) 2025.02.03.xlsx'.
n. rows:	 803
Columns:	 ['FULL_NAME (LFM)', 'CHOSEN FIRST NAME', 'PHONE', 'EMAIL_ADDRESS', 'ADDRESS_LINE1', 'ADDRESS_LINE2', 'TOWN/CITY', 'ST', 'ZIP', 'Program/Field of Study']
Column(s) with null values:
               % null  total null
ADDRESS_LINE2   51.31         412
ST               0.37           3

Parsed employer list date as '2025.02.03'.

Parsing program/field of study data...
Parsed 23 different programs/departments and 9 different degree types.

Parsed full name -> First Last Middle.

Combined address column created.

Finished parsing this employer BU list. Please check the output for errors before using it.
Wrote to file 'data/BU List Employer 2025.02.03 made 2025.02.06_10.57.53.csv'
```

```bash
python parse_employer_bu.py --infile="data/GOLD Membership (Fall 2024) - supplemented 2024.10.31.xlsx" --program_col="Program/Field of Study" --lfm_cols LAST FIRST MIDDLE
```

```bash
python parse_employer_bu.py --infile="data/GOLD Membership (Fall 2024) 2024.10.15.xlsx" --program_col="PROGRAM/FIELD OF STUDY" --lfm_cols LAST FIRST MIDDLE
```
