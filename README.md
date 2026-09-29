# listwork-scripts

Scripts for the recurring listwork behind GOLD's membership database in Broadstripes.

The end-to-end process these fit into — including everything that happens in the
Broadstripes web app — is written up in
[`process_employer_list_doc.md`](process_employer_list_doc.md). **Start there if
you are new to this job.** This README is the reference for the two scripts.

## Setup

Requires `python>=3.11` plus `pandas` and `openpyxl`.

If you don't already have a python environment with those:

```bash
# Using conda/mamba/miniforge
mamba create -n env-listwork python=3.13 pandas openpyxl pytest
mamba activate env-listwork

# Or using plain python + venv
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements-dev.txt
```

Then clone this repo and run the scripts from its root directory:

```bash
git clone https://github.com/covertg/listwork-scripts.git
cd listwork-scripts
```

Everything in `data/` is gitignored, because BU lists contain workers' personal information. Keep it that way. For testing, it can be helpful to pull old employer BU lists from our internal Google Drive folder.

## `parse_employer_bu.py`

Turns a BU list from Dartmouth (`.xlsx`) into a CSV that's ready to import into
Broadstripes.

### Example usage (peeking at the data)

Dartmouth renames the columns fairly often, so start by seeing what you've got:

```bash
python parse_employer_bu.py -i './data/2026.07.28 TO GOLD Membership 26X.xlsx' --show_columns
```

That prints the column names and the first few rows, then exits. Use it to work
out what to pass for `--program_col`, `--fullname_col`, and `--address_cols`.

### Example usage (parsing)

```bash
python parse_employer_bu.py \
    -i './data/2026.07.28 TO GOLD Membership 26X.xlsx' \
    --program_col "Program1 Code" \
    --fullname_col "LFM Name Formatted" \
    --address_cols "LO Street1" "LO Street2" "LO City" "LO State" "LO Zip"
```

Those were the right arguments for the 26W, 26S, and 26X lists.

Output:

```
Loaded file 'data/2026.07.28 TO GOLD Membership 26X.xlsx'.
n. rows:	 649
Columns:	 ['LFM Name Formatted', 'Chosen-Pref-Legal First Name', 'DND Email', 'LO Street1', 'LO Street2', 'LO City', 'LO State', 'LO Zip', 'MP Phone For Address', 'Program1 Code']
Column(s) with null values:
                      % null  total null
LO Street1              1.39           9
LO Street2             38.06         247
...

Parsed employer list date as '2026.07.28'.

Parsing program/field of study data...
Parsed 21 different programs/departments and 7 different degree types.

Parsed full name -> First Last Middle.

Combined address column created.

Checking data quality...
No data quality issues found.

Finished parsing this employer BU list. Please check the output for errors before using it.
Wrote to file 'data/BU List Employer 2026.07.28 made 2026.09.04_22.48.42.csv'
```

By default the CSV is written next to the input file (so, normally, `data/`)
with the list identifier and a timestamp in its name. Pass `-o` to choose a
different path. `python parse_employer_bu.py --help` documents every option.

### When you might need to intervene

There are some data issues that the script will complain about, and you may need to address:

**New program codes** — the most common one (happens most terms). The script 
stops with a list of the codes it doesn't recognize and the Excel row numbers of
the workers under each. Work out what each new code means, add it to `program_mapping.toml`, and re-run with the same arguments.

**Suspicious rows** — printed as `!! WARNING`. As of writing, suspicious rows get flagged if:

- *Rows sharing a name.* Sometimes the same worker exported twice; sometimes two
  real people. Compare their other info to tell them apart.
- *Shifted address columns.* Dartmouth's export sometimes pushes a row's
  city/state/zip/phone one column to the right, which silently turns
  someone's phone number into a zip code.
- *A US state with a malformed zip.* Such as a dropped leading zero (`3755`) or an extra digit (`037555`).
- *States/zips that aren't US-shaped.* Usually just an international address.

Some of these warnings might not be issues, other ones you will want to fix in the `xlsx`. If you edit the `xlsx`, save the edited version as a new file, so we don't lose the original data.

**Missing data** — see the "columns with null values" portion of the script output, and use your eyes to do a visual check through the data.

### Tests

If modifying this script, check and potentially add tests. See
[`tests/parse_employer_bu/README.md`](tests/parse_employer_bu/README.md).

```bash
pytest tests/
```

## `check_skipped_imports.py`

After the first Broadstripes import (which matches workers by email), Broadstripes
gives us a CSV of "skips": the rows it couldn't match. Most are new workers, but
some are existing workers whose email changed, often along with their name. This
script compares every skip against a full export of Broadstripes and lists
possible matches, so we can fix those workers' emails instead of creating
duplicates.

A Broadstripes record counts as a possible match if it has:

- **the same name** or **the same phone number** (a *strong* match: check these
  carefully), or
- **a similar name** based on fuzzy name matching (a *weaker* match: most of these are just different people
  who share a first name). `--threshold` (default 0.75) sets how similar.

### Example usage

The defaults expect the skips CSV straight from Broadstripes (its columns come
from `parse_employer_bu.py`) and a full export using the "Contact and Degree
Info" layout, so normally you just pass the two files:

```bash
python check_skipped_imports.py \
    -s './data/2026-04-29 data-import-SKIPS-0d7dc05d-34ec-4a77-a28a-91f374f3d298.csv' \
    -b './data/2026.09.29 Contact and Degree Info.csv'
```

Output (names are made up):

```
Loaded skips file 'data/2026-04-29 data-import-SKIPS-....csv' (65 rows).
Loaded Broadstripes file 'data/2026.09.29 Contact and Degree Info.csv' (2587 rows).

Compared 65 skipped rows against 2587 Broadstripes records.
  44 skips have no possible match (likely new workers).
  21 skips have possible matches:
    2 strong (same name or same phone number). Check these carefully.
    19 weaker (similar name only). Most of these are different people.

=== Strong possible matches ===

                                  Name          Phone                         Email Employer Broadstripes ID Name similarity
Skip (CSV row 42)    Rivera, Jordan A. (603) 555-0142 Jordan.A.Rivera.TH@dartmouth.edu     ENGS
Match: same phone     Chen, Jordan A.     6035550142  Jordan.A.Chen.TH@dartmouth.edu     ENGS         ABC-123            0.62
...

=== Weaker possible matches (similar name only) ===
...

Wrote these possible matches to 'data/2026-04-29 data-import-SKIPS-... - possible duplicates.csv'
```

The same tables are saved to a CSV next to the skips file (or pass `-o`), which
is easier to read in a spreadsheet. "CSV row" is the row number in the skips file
as a spreadsheet shows it.

If Broadstripes or Dartmouth renames a column, the script stops and lists the
file's actual columns. Use the `--*_name_cols` and `--*_info_cols` options to
point it at the new names; `python check_skipped_imports.py --help` documents
every option.

### Tests

```bash
pytest tests/check_skipped_imports/
```

If we wanted to improve the matching further, address is another promising signal (the Broadstripes export would need to include it).
