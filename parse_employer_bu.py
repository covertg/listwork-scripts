"""Parse a Dartmouth employer BU list into a CSV that is ready to import into Broadstripes.

Run `python parse_employer_bu.py --help` for usage, and see README.md for a
worked example. The overall listwork process this fits into is documented in
process_employer_list_doc.md.
"""

try:
    import pandas as pd
except ImportError as e:
    raise ImportError(
        "This script requires the python library pandas. Please install it to your environment."
    ) from e
try:
    import openpyxl  # noqa: F401
except ImportError as e:
    raise ImportError(
        "This script requires the python library openpyxl. Please install it to your environment."
    ) from e
import argparse
import datetime
from pathlib import Path
from pprint import pprint
import re
import sys
import tomllib


# Default address column names used by make_address_combined
DEFAULT_ADDRESS_COLS = ("ADDRESS_LINE1", "ADDRESS_LINE2", "TOWN/CITY", "ST", "ZIP")

# A middle name in these lists is always a single initial followed by a period,
# e.g. the "A." in "Smith, John A.". Matching that exactly (rather than any token
# ending in ".") keeps suffixes like "Jr." out of the Middle column.
MIDDLE_INITIAL_PATTERN = re.compile(r"^[^\W\d_]\.$", re.UNICODE)

# What a well-formed US state and zip look like in Dartmouth's export.
STATE_PATTERN = re.compile(r"^[A-Za-z]{2}$")
ZIP_PATTERN = re.compile(r"^\d{5}(-\d{4})?$")


def _warn(message: str) -> None:
    """Print a warning that is hard to miss when skimming the script's output."""
    print(f"\n!! WARNING: {message}")


def check_columns(df: pd.DataFrame, wanted: dict[str, str]) -> None:
    """Check up front that every column we were asked to use actually exists.

    `wanted` maps an argument name (e.g. "--program_col") to a column name. We
    check them all at once, before any parsing work, so that a typo in one
    argument is reported immediately and alongside the real column names,
    rather than surfacing as a pandas KeyError partway through the run.
    """
    missing = {arg: col for arg, col in wanted.items() if col not in df.columns}
    if not missing:
        return
    print("\nError: some requested column(s) are not in the input file:")
    for arg, col in missing.items():
        print(f"  {arg}: '{col}'")
    print("The input file's actual columns are:")
    for col in df.columns:
        print(f"  '{col}'")
    print(
        "Dartmouth renames these columns fairly often. Fix the argument(s) above to "
        "match, or run with --show_columns to peek at the file first."
    )
    raise ValueError(f"Column(s) not found in input file: {sorted(missing.values())}")


def load_list(infile: Path) -> pd.DataFrame:
    """Load list and check that the expected columns are there"""
    if not infile.exists():
        raise FileNotFoundError(f"Input file '{infile}' does not exist")
    # Warn if the workbook has more than one sheet: pd.read_excel silently reads
    # only the first one, and we'd rather not silently drop half a BU list.
    sheets = pd.ExcelFile(infile).sheet_names
    if len(sheets) > 1:
        _warn(
            f"'{infile.name}' has {len(sheets)} sheets {sheets}. Only the first "
            f"('{sheets[0]}') will be read. Check that this is what you want."
        )
    # Load file
    # It's important to use dtype=object or str, otherwise zip code may be cast to float
    df = pd.read_excel(infile, dtype=str)
    print(f"Loaded file '{infile}'.")
    print(f"n. rows:\t {df.shape[0]}")
    print(f"Columns:\t {df.columns.tolist()}")
    # Strip unecessary whitespace from all columns (sometimes name data has this issue)
    # read_excel(dtype=str) always gives us object columns, but guard anyway so
    # that this function is also safe to call on a hand-built dataframe.
    df = df.apply(lambda col: col.str.strip() if col.dtype == object else col)
    # Monitor data missingness
    blanks = df == ""
    if blanks.any().any():
        print(
            "\nError: some cells in the input contain only whitespace rather than "
            "being truly empty. We can't tell whether these are meant to be blank "
            "or are lost data, so please check them in the .xlsx."
        )
        print("Row numbers below are as shown in Excel (header is row 1):")
        for col in blanks.columns[blanks.any(axis=0)]:
            rows = [i + 2 for i in blanks.index[blanks[col]]]
            print(f"  column '{col}': row(s) {rows}")
        raise RuntimeError(
            "Input data has empty (non-nan) cells, please check data quality"
        )
    nas = df.isna()
    nas = nas.loc[:, nas.any(axis=0)]
    if nas.any().any():
        nas = pd.concat(
            [(nas.mean() * 100).round(2), nas.sum()],
            axis=1,
            keys=["% null", "total null"],
        )
        print("Column(s) with null values:")
        print(nas)
    return df


def extract_date(text: str) -> str | None:
    """Searches for a date in the form YYYY.MM.DD in a string."""
    # Pattern explanation:
    # \d{4} - exactly 4 digits for year
    # \. - literal dot
    # \d{2} - exactly 2 digits for month
    # \. - literal dot
    # \d{2} - exactly 2 digits for day
    pattern = r"(\d{4})\.(\d{2})\.(\d{2})"
    match = re.search(pattern, text)
    if match:
        year, month, day = match.groups()
        # Check date validity
        try:
            _ = datetime.date(year=int(year), month=int(month), day=int(day))
        except ValueError:
            print(f"Invalid date in input: '{text}'")
            return None
        return f"{year}.{month}.{day}"
    return None


def get_list_identifier(infile: Path) -> str:
    """Determines unique identifier for this generated list, currently 'BU List Employer YYYY.MM.DD'"""
    date_received = extract_date(infile.name)
    if not date_received:
        raise ValueError(
            f"Input BU file needs to have the date received somewhere in its filename in the format YYYY.MM.DD. We could not parse a valid date from the filename {infile.name}."
        )
    print(f"\nParsed employer list date as '{date_received}'.")
    return f"BU List Employer {date_received}"


def convert_program_mapping(
    df: pd.DataFrame, program_col: str, program_mapping_file: Path
) -> pd.DataFrame:
    print("\nParsing program/field of study data...")
    if program_col not in df.columns:
        raise ValueError(
            f"Input for program_col of '{program_col}' does not exist in the dataframe, please check its name"
        )
    # Load program mapping file (this is our hand-made file)
    if not program_mapping_file.exists():
        raise FileNotFoundError(
            f"Could not find the program mapping file '{program_mapping_file}'"
        )
    with open(program_mapping_file, "rb") as f:
        program_mapping = tomllib.load(f)
    # Check the Dartmouth list for missing data in program/field of study
    if df[program_col].hasnans:
        raise RuntimeError(
            f"Original data for program/field of study (column '{program_col}') has null values"
        )
    # Initial cleanup, sometimes the Dartmouth data has a redundant "PROGRAM" text that we don't need
    df[program_col] = df[program_col].str.replace(" PROGRAM", "", case=False)
    # Check if there are any new programs that we haven't seen before
    new_programs = set(df[program_col]) - set(program_mapping.keys())
    if new_programs:
        print(
            "Error: This file has new entries in program/field of study that we haven't seen before. This often happens with each new BU list. Please interpret the new entries and add them to the program mapping .toml file. Unrecognized entries:"
        )
        # Show how many workers each unknown code covers and where they are in
        # the file (Excel row numbers, header being row 1). Looking those workers
        # up is usually what lets you figure out what a new code means.
        for code in sorted(new_programs):
            rows = [i + 2 for i in df.index[df[program_col] == code]]
            print(f"  '{code}' — {len(rows)} worker(s), Excel row(s) {rows}")
        raise RuntimeError("Unrecognized program/field of study data")
    # If everything looks good, then map program/field of study to new Employer and Degree columns
    df[["Employer", "Degree"]] = df.apply(
        lambda row: program_mapping[row[program_col]],
        axis=1,
        result_type="expand",
    )
    print(
        f"Parsed {df['Employer'].unique().size} different programs/departments and {df['Degree'].unique().size} different degree types."
    )
    return df


def parse_fullnames(df: pd.DataFrame, fullname_col: str) -> pd.DataFrame:
    """Split names in the "fullname" format to separate last, first, and middle columns.

    In some BU lists, Dartmouth provides names only in the format of:
    [Last name], [First name(s)] [Middle initial (optional)].
    """
    if fullname_col not in df.columns:
        raise ValueError(
            f"Input for fullname_col of '{fullname_col}' does not exist in the dataframe, please check its name"
        )
    last_names = []
    first_names = []
    middle_names = []
    for fullname in df[fullname_col]:
        if not isinstance(fullname, str) or not fullname.strip():
            raise ValueError(f"Invalid name value: {fullname}")
        # Split on first comma to get last name and rest
        parts = fullname.split(",")
        if len(parts) != 2:
            raise ValueError(f"Name must contain exactly one comma: {fullname}")
        last, rest = parts
        last = last.strip()
        rest = rest.strip()
        # Split the "rest" on spaces
        parts = rest.split()
        if MIDDLE_INITIAL_PATTERN.match(parts[-1]):  # Is there a middle initial?
            middle = parts[-1]
            first = " ".join(parts[:-1])
        else:  # Sometimes there is not, and we treat the whole "rest" as first name
            first = " ".join(parts)
            middle = ""
        last_names.append(last)
        first_names.append(first)
        middle_names.append(middle)
    # And we're done, add to the input dataframe
    df["Last"] = last_names
    df["First"] = first_names
    df["Middle"] = middle_names
    # Check that no name data was lost from original (the operation is invertible)
    fullnames_reconstructed = (
        df["Last"] + ", " + df["First"] + " " + df["Middle"]
    ).str.strip()
    nonmatches = fullnames_reconstructed != (df[fullname_col])
    if nonmatches.sum() != 0:
        print(
            "Error: We received some name(s) in an unexpected format. Please check the following entries:"
        )
        pprint(df.loc[nonmatches, fullname_col])
        raise ValueError("Full name data in unexpected format")
    print("\nParsed full name -> First Last Middle.")
    return df


def check_names_lfm(df: pd.DataFrame, lastc: str, firstc: str, middlec: str):
    """Simple check that we have no missing data for Last and First names."""
    if df[lastc].isna().any():
        raise ValueError(f"Missing data in Last name column '{lastc}'")
    if df[firstc].isna().any():
        raise ValueError(f"Missing data in First name column '{firstc}'")
    if df[middlec].isna().all():
        print(f"\nWarning: 100% missing data in Middle name column '{middlec}'")


def str_combine(*e) -> str:
    """String-combines a list of address elements.

    Empty strings and nan-like elements are omitted, and elements are
    strip()ed of whitespace and trailing commas (because some employer
    data has unnecessary commas). The output is separated by spaces.
    """
    strs = [str(s).strip().rstrip(",") for s in e if (not pd.isna(s) and s)]
    return " ".join(strs)


def make_address_combined(
    df: pd.DataFrame,
    l1c: str,
    l2c: str,
    cityc: str,
    statec: str,
    zipc: str,
) -> pd.DataFrame:
    addrs = []
    for _, row in df.iterrows():
        line1, line2, town, st, zipcode = row.loc[[l1c, l2c, cityc, statec, zipc]]
        # Combine strings before the comma
        addr = str_combine(line1, line2)
        # Combine strings after the comma. Note that we test the post-comma part
        # via str_combine rather than the raw values, because a NaN is truthy in
        # python and would otherwise leave us with a trailing ", ".
        rest = str_combine(town, st, zipcode)
        if rest:
            if addr:
                addr += ", "
            addr += rest
        addrs.append(addr)
    df["Address Combined"] = addrs
    print("\nCombined address column created.")
    return df


def _malformed(col: pd.Series, pattern: re.Pattern) -> pd.Series:
    """Boolean mask of values that are present but don't match `pattern`.

    Missing values are never malformed. Written with a plain map rather than
    `.str.match` so that an all-empty column (which pandas hands back as
    float64, with no .str accessor) doesn't blow up.
    """
    return col.map(
        lambda v: False if pd.isna(v) else pattern.match(str(v)) is None
    ).astype(bool)


def check_data_quality(
    df: pd.DataFrame,
    name_cols: list[str],
    cityc: str,
    statec: str,
    zipc: str,
) -> int:
    """Look for the kinds of bad rows we have actually seen in employer lists.

    This never stops the run: some of what it flags is legitimate (international
    addresses have no US state or zip), and other things that it flags would be too
    fidgety/unecessary to try to fix it in code. It prints what it finds so that you can
    eyeball the rows, fix them in the xlsx as needed, and re-run if needed. Returns the
    number of rows flagged.

    Row numbers are printed as they appear in Excel (header is row 1) so that
    they can be looked up directly in the source file.
    """
    print("\nChecking data quality...")
    flagged: set[int] = set()

    def _excel_rows(index) -> list[int]:
        return [i + 2 for i in index]

    # 1. Rows that share a name. In the 26X list, one worker appeared twice,
    #    with the second copy's address columns shifted. Importing both makes a
    #    duplicate (or overwrites good data with bad).
    dupes = df[df.duplicated(subset=list(name_cols), keep=False)]
    if not dupes.empty:
        flagged.update(dupes.index)
        _warn(
            f"{len(dupes)} row(s) share a name with another row. These are usually "
            "an export glitch (the same worker listed twice), but can occasionally "
            "be two real people. Excel row(s): "
            f"{_excel_rows(dupes.index)}"
        )
        print(dupes[list(name_cols)].to_string())

    # 2. Shifted address columns. We have seen Dartmouth exports where, for some
    #    rows, city lands in the state column, state in the zip column, and the
    #    zip in whatever column comes next (often the phone number).
    shifted = df[df[cityc].isna() & df[statec].notna()]
    if not shifted.empty:
        flagged.update(shifted.index)
        _warn(
            f"{len(shifted)} row(s) have an empty city but a non-empty state. This is "
            "the signature of address columns being shifted one to the right in "
            "Dartmouth's export: the city sits in the state column, the state in "
            "the zip column, and the ZIP CODE lands in the phone column. Check the "
            "phone numbers for these rows too — Broadstripes will happily import a "
            "zip code as someone's phone number. Excel row(s): "
            f"{_excel_rows(shifted.index)}"
        )
        print(shifted[[*name_cols, cityc, statec, zipc]].to_string())

    bad_state = _malformed(df[statec], STATE_PATTERN)
    bad_zip = _malformed(df[zipc], ZIP_PATTERN)

    # 3. A well-formed US state next to a malformed zip. This is almost always a
    #    typo in the worker's own record rather than an export problem, and it
    #    turns up in nearly every list: a lost leading zero ("3755") or a stray
    #    extra digit ("037555"). Broadstripes will import it verbatim.
    us_bad_zip = df.index[df[statec].notna() & ~bad_state & bad_zip].difference(
        shifted.index
    )
    if len(us_bad_zip):
        flagged.update(us_bad_zip)
        _warn(
            f"{len(us_bad_zip)} row(s) have a US state but a zip that isn't 5 digits "
            "(or 5+4). Usually a typo in the worker's record — a dropped leading "
            "zero or an extra digit. Worth correcting before import. Excel row(s): "
            f"{_excel_rows(us_bad_zip)}"
        )
        print(df.loc[us_bad_zip, [*name_cols, cityc, statec, zipc]].to_string())

    # 4. Anything else that isn't US-shaped. International addresses trip this
    #    legitimately, so it's informational.
    odd = df.index[bad_state | bad_zip].difference(shifted.index).difference(us_bad_zip)
    if len(odd):
        flagged.update(odd)
        _warn(
            f"{len(odd)} row(s) have a state or zip that doesn't look like a US "
            "state/zip. International addresses do this legitimately — just "
            f"eyeball them. Excel row(s): {_excel_rows(odd)}"
        )
        print(df.loc[odd, [*name_cols, cityc, statec, zipc]].to_string())

    if flagged:
        print(
            f"\n{len(flagged)} row(s) flagged above. None of this stops the parse. "
            "Fix anything genuinely wrong in the source .xlsx and re-run, or note "
            "it and move on."
        )
    else:
        print("No data quality issues found.")
    return len(flagged)


def parse_dartmouth_bu(
    infile: Path,
    program_col: str,
    name_cols: list[str],
    program_mapping_file: Path,
    outfile: Path | None,
    address_cols: tuple[str, str, str, str, str],
    write: bool = True,
) -> pd.DataFrame:
    with pd.option_context("mode.copy_on_write", True):
        df = load_list(infile=infile)

        # Check every column we were asked to use before doing any work, so that
        # a mistyped argument fails fast with a useful message.
        wanted = {"--program_col": program_col}
        if len(name_cols) == 1:
            wanted["--fullname_col"] = name_cols[0]
        else:
            for arg, col in zip(("last", "first", "middle"), name_cols):
                wanted[f"--lfm_cols ({arg})"] = col
        for arg, col in zip(("line1", "line2", "city", "state", "zip"), address_cols):
            wanted[f"--address_cols ({arg})"] = col
        check_columns(df, wanted)

        # Parse received date and make BU list column
        list_name = get_list_identifier(infile)
        df[list_name] = True

        # Parse program mapping
        df = convert_program_mapping(
            df=df, program_col=program_col, program_mapping_file=program_mapping_file
        )

        # Check and parse names
        if len(name_cols) == 1:
            # Convert fullname column to three columns, Last First Middle
            df = parse_fullnames(df, name_cols[0])
        elif len(name_cols) == 3:
            check_names_lfm(df, *name_cols)
        else:
            raise ValueError(
                f"Invalid number of name columns provided for name_cols '{name_cols}'. Expected 1 (fullname) or 3 (LFM)."
            )

        # Create combined address column
        df = make_address_combined(df, *address_cols)

        # Flag suspicious rows (advisory only/never stops the run)
        _, _, cityc, statec, zipc = address_cols
        n_flagged = check_data_quality(
            df, name_cols=name_cols, cityc=cityc, statec=statec, zipc=zipc
        )

        print(
            "\nFinished parsing this employer BU list. Please check the output for errors before using it."
        )
        if n_flagged:
            print(f"Remember to look at the {n_flagged} flagged row(s) above.")
        if write:
            if not outfile:
                fname = f"{list_name} made {datetime.datetime.now().strftime('%Y.%m.%d_%H.%M.%S')}.csv"
                # Default to writing next to the input file, which is normally
                # the repo's data/ directory. This works no matter which
                # directory you happen to run the script from.
                outfile = infile.parent / fname
            outfile.parent.mkdir(parents=True, exist_ok=True)
            df.to_csv(outfile, index=False)
            print(f"Wrote to file '{outfile}'")
        return df


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Parse a Dartmouth BU list into a CSV file, cleaned and formatted to use with Broadstripes.",
        epilog=(
            "Not sure what the columns are called in this term's file? Run with just "
            "-i FILE --show_columns to print them."
        ),
    )
    parser.add_argument(
        "-i",
        "--infile",
        type=Path,
        help="Path to the input file (.xlsx from employer). The filename must include the date that we received it in the format YYYY.MM.DD",
        required=True,
    )
    parser.add_argument(
        "-o",
        "--outfile",
        type=Path,
        help="Path to the output file (.csv). Optional. By default, the output file will be written next to the input file (normally the data/ directory) with an informative name and a timestamp.",
        default=None,
        required=False,
    )
    parser.add_argument(
        "--show_columns",
        action="store_true",
        help="Print the input file's column names and first few rows, then exit without parsing. Use this first when you get a new BU list, to work out what to pass for --program_col, --fullname_col, and --address_cols.",
    )
    parser.add_argument(
        "--program_col",
        type=str,
        help="Name of the column containing the program/field of study. Required (unless --show_columns). This column name seems to vary pretty frequently by term, so you will need to identify it by peeking at the input file.",
        required=False,
    )
    parser.add_argument(
        "--fullname_col",
        type=str,
        help="Name of the column containing the full name. Exactly one of fullname_col or lfm_cols must be specified. All of recent BU lists from Dartmouth have used a unified column for the full name, so this is probably the argument you want for a new BU list.",
        required=False,
    )
    parser.add_argument(
        "--lfm_cols",
        nargs=3,
        metavar=("LAST", "FIRST", "MIDDLE"),
        help="(Legacy) Name of the columns containing the last name, first name, and middle initial, separated by spaces. Exactly one of lfm_cols or fullname_col must be specified. Dartmouth's original BU lists separated names into columns, but more recent lists have all used a unified column, so you probably don't want this argument for a new BU list.",
        required=False,
    )
    parser.add_argument(
        "--program_mapping_file",
        type=Path,
        help="Path to the program mapping file (.toml that we maintain). Defaults to ./program_mapping.toml",
        default="program_mapping.toml",
        required=False,
    )
    parser.add_argument(
        "--address_cols",
        nargs=5,
        metavar=("LINE1", "LINE2", "CITY", "STATE", "ZIP"),
        help="Names of the address columns: line1, line2, city, state, and zip (in that order).",
        default=list(DEFAULT_ADDRESS_COLS),
        required=False,
    )
    return parser


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.show_columns:
        df = load_list(args.infile)
        print("\nFirst rows:")
        with pd.option_context("display.max_columns", None, "display.width", 250):
            print(df.head())
        print(
            "\nUse the column names above for --program_col, --fullname_col, and "
            "--address_cols. See README.md for a full example command."
        )
        return

    if not args.program_col:
        parser.error(
            "--program_col is required (or use --show_columns to peek at the file first)."
        )
    # Error if lfm_cols and fullname_col are both or neither specified
    if bool(args.fullname_col) == bool(args.lfm_cols):
        parser.error(
            "Please specify exactly one of --fullname_col or --lfm_cols so we know "
            "which column(s) to use for names. Recent Dartmouth lists use a single "
            "full-name column, so --fullname_col is probably the one you want."
        )
    name_cols = [args.fullname_col] if args.fullname_col else args.lfm_cols

    _ = parse_dartmouth_bu(
        infile=args.infile,
        program_col=args.program_col,
        name_cols=name_cols,
        program_mapping_file=args.program_mapping_file,
        outfile=args.outfile,
        address_cols=tuple(args.address_cols),
        write=True,
    )


if __name__ == "__main__":
    main(sys.argv[1:])
