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
import tomllib

pd.options.mode.copy_on_write = True


def load_list(infile: Path) -> pd.DataFrame:
    """Load list and check that the expected columns are there"""
    if not infile.exists():
        raise FileNotFoundError(f"Input file '{infile}' does not exist")
    # Load file
    # It's important to use dtype=object or str, otherwise zip code may be cast to float
    df = pd.read_excel(infile, dtype=str)
    print(f"Loaded file '{infile}'.")
    print(f"n. rows:\t {df.shape[0]}")
    print(f"Columns:\t {df.columns.tolist()}")
    # Strip unecessary whitespace from all columns (sometimes name data has this issue)
    df = df.apply(lambda col: col.str.strip())
    # Monitor data missingness
    if (df == "").any().any():
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
            "Input BU file needs to have the date received somewhere in its filename in the format YYYY.MM.DD. We could not parse a valid date."
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
        pprint(new_programs)
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
    if fullname_col and fullname_col not in df.columns:
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
        if parts[-1].endswith("."):  # Is there a middle name?
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
    df,
    l1c="ADDRESS_LINE1",
    l2c="ADDRESS_LINE2",
    cityc="TOWN/CITY",
    statec="ST",
    zipc="ZIP",
):
    addrs = []
    for _, row in df.iterrows():
        line1, line2, town, st, zip = row.loc[[l1c, l2c, cityc, statec, zipc]]
        # Combine strings before the comma
        addr = str_combine(line1, line2)
        if town or st or zip:
            if addr:
                addr += ", "
            # Combine strings after the comma
            addr += str_combine(town, st, zip)
        addrs.append(addr)
    df["Address Combined"] = addrs
    print("\nCombined address column created.")
    return df


def parse_dartmouth_bu(
    infile: Path,
    program_col: str,
    name_cols: list[str],
    program_mapping_file: Path,
    outfile: Path | None,
    write=True,
):
    df = load_list(infile=infile)
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
    df = make_address_combined(df)

    print(
        "\nFinished parsing this employer BU list. Please check the output for errors before using it."
    )
    if write:
        if not outfile:
            fname = f"{list_name} made {datetime.datetime.now().strftime('%Y.%m.%d_%H.%M.%S')}.csv"
            outfile = Path("data") / fname
        df.to_csv(outfile, index=False)
        print(f"Wrote to file '{outfile}'")
    return df


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Parse a Dartmouth BU list into a CSV file, cleaned and formatted to use with Broadstripes."
    )
    parser.add_argument(
        "--infile",
        type=Path,
        help="Path to the input file (.xlsx from employer)",
        required=True,
    )
    parser.add_argument(
        "--program_col",
        type=str,
        help="Name of the column containing the program/field of study",
        required=True,
    )
    parser.add_argument(
        "--fullname_col",
        type=str,
        help="Name of the column containing the full name. Only fullname_Col or lfm_cols can be specified.",
        required=False,
    )
    parser.add_argument(
        "--lfm_cols",
        nargs=3,
        help="Name of the columns containing the last name, first name, and middle initial, separated by spaces. Only lfm_cols or fullname_col can be specified.",
        required=False,
    )
    parser.add_argument(
        "--program_mapping_file",
        type=Path,
        help="Path to the program mapping file (.toml from us)",
        default="program_mapping.toml",
        required=False,
    )
    parser.add_argument(
        "--outfile",
        type=Path,
        help="Path to the output file (.csv). Optional. By default, the output file will be in the data/ directory and have an informative name with a timestamp.",
        default=None,
        required=False,
    )
    args = parser.parse_args()

    # Error if lfm_cols and fullname_col are both or neither specified
    if (args.fullname_col and args.lfm_cols) or not (args.fullname_col or args.lfm_cols):
        raise ValueError(
            "Please specific exactly one of 'fullname_col' or 'lfm_cols' so we know which column(s) to use for names."
        )
    name_cols = [args.fullname_col] if args.fullname_col else args.lfm_cols

    _ = parse_dartmouth_bu(
        infile=args.infile,
        program_col=args.program_col,
        name_cols=name_cols,
        program_mapping_file=args.program_mapping_file,
        outfile=args.outfile,
        write=True,
    )
