try:
    import pandas as pd
except ImportError as e:
    raise ImportError(
        "This script requires the python library pandas. Please install it to your environment."
    ) from e

import argparse
from difflib import SequenceMatcher
from pathlib import Path

pd.options.mode.copy_on_write = True

DEMO_FIELDS = ["Phone", "Address", "Email", "Employer", "Degree"]


def find_potential_matches(
    names1: list[str], names2: list[str], threshold: float = 0.75
) -> list[dict]:
    """Compare two lists of names to find exact and fuzzy matches.

    Parameters:
        names1: Names from the skips file (indexed by position)
        names2: Names from the existing Broadstripes file (indexed by position)
        threshold: Float between 0 and 1 for fuzzy matching similarity threshold

    Returns:
        List of match groups sorted by best similarity descending. Each group is:
        - skip_idx: index into names1
        - skip_name: cleaned name
        - best_similarity: highest similarity score among matches
        - matches: list of (existing_idx, existing_name, match_type, similarity)
    """
    results = []
    cleaned1 = [
        (i, str(n).lower().strip()) for i, n in enumerate(names1) if pd.notna(n)
    ]
    cleaned2 = [
        (i, str(n).lower().strip()) for i, n in enumerate(names2) if pd.notna(n)
    ]

    # Build lookup for exact matches (one name can map to multiple rows)
    name2_by_value: dict[str, list[int]] = {}
    for idx, name in cleaned2:
        name2_by_value.setdefault(name, []).append(idx)

    for skip_idx, name in cleaned1:
        matches = []
        if name in name2_by_value:
            for existing_idx in name2_by_value[name]:
                matches.append((existing_idx, name, "exact", 1.0))
        else:
            for existing_idx, existing_name in cleaned2:
                similarity = SequenceMatcher(
                    None, name, existing_name, autojunk=False
                ).ratio()
                if similarity >= threshold:
                    matches.append((existing_idx, existing_name, "fuzzy", similarity))

        if matches:
            matches.sort(key=lambda x: x[3], reverse=True)
            results.append(
                {
                    "skip_idx": skip_idx,
                    "skip_name": name,
                    "best_similarity": matches[0][3],
                    "matches": matches,
                }
            )

    results.sort(key=lambda x: x["best_similarity"], reverse=True)
    return results


def build_match_table(
    match_group: dict,
    skip_df: pd.DataFrame,
    existing_df: pd.DataFrame,
    skip_demo_cols: list[str],
    existing_demo_cols: list[str],
) -> pd.DataFrame:
    """Build a DataFrame showing a skip entry and its possible matches with demographics.

    Parameters:
        match_group: A match group dict from find_potential_matches
        skip_df: Full DataFrame of skipped entries
        existing_df: Full DataFrame of existing Broadstripes entries
        skip_demo_cols: Demographic column names in skip_df
            (ordered: Phone, Address, Email, Employer, Degree)
        existing_demo_cols: Demographic column names in existing_df
            (ordered: Phone, Address, Email, Employer, Degree)

    Returns:
        DataFrame with columns: [Role, Name, Phone, Address, Email, Employer, Degree, Similarity]
    """
    rows = []
    skip_idx = match_group["skip_idx"]

    # Entry row (from skips file)
    entry_data = skip_df.iloc[skip_idx]
    row = {"": "Entry", "Name": match_group["skip_name"], "Similarity": "-"}
    for field, col in zip(DEMO_FIELDS, skip_demo_cols):
        val = entry_data.get(col, "")
        row[field] = val if pd.notna(val) else ""
    rows.append(row)

    # Possible match rows (from existing Broadstripes file)
    for existing_idx, existing_name, match_type, similarity in match_group["matches"]:
        existing_data = existing_df.iloc[existing_idx]
        row = {
            "": f"Possible match ({match_type})",
            "Name": existing_name,
            "Similarity": f"{similarity:.4f}",
        }
        for field, col in zip(DEMO_FIELDS, existing_demo_cols):
            val = existing_data.get(col, "")
            row[field] = val if pd.notna(val) else ""
        rows.append(row)

    columns = ["", "Name"] + DEMO_FIELDS + ["Similarity"]
    return pd.DataFrame(rows, columns=columns)


def load_and_validate_csv(
    file: Path, required_cols: list[str], file_description: str
) -> pd.DataFrame:
    """Load a CSV file and validate required columns exist.

    Parameters:
        file: Path to CSV file
        required_cols: All column names that must exist (name cols + demographic cols)
        file_description: Description of file for error messages

    Returns:
        Loaded DataFrame
    """
    df = pd.read_csv(file, dtype=str)

    missing_cols = set(required_cols) - set(df.columns)
    if missing_cols:
        raise ValueError(
            f"Missing required columns in {file_description}: {missing_cols}"
        )

    print(f"Loaded {file_description} from '{file}'")
    print(f"n. rows:\t {df.shape[0]}")
    return df


def combine_name_parts(df, columns):
    """
    Combines name parts from specified columns, handling NaN values gracefully.

    Parameters:
        df: pandas DataFrame containing name columns
        columns: list of column names [last_name, first_name, middle_name]

    Returns:
        list of combined names, with proper handling of NaN values
    """

    def clean_name_part(x):
        return str(x) if pd.notna(x) else ""

    combined_names = []
    for parts in zip(*[df[col] for col in columns]):
        cleaned_parts = [clean_name_part(part) for part in parts]
        last, first, middle = cleaned_parts
        full_name = f"{last}, {first} {middle}".strip()
        if not full_name:
            raise ValueError(f"Invalid name: {parts}")
        combined_names.append(full_name)
    return combined_names


def check_skipped_imports(
    all_broadstripes: Path,
    skipped_entries: Path,
    all_bs_cols: list[str],
    skipped_cols: list[str],
    all_bs_demo_cols: list[str],
    skipped_demo_cols: list[str],
    similarity_threshold: float,
    outfile: Path | None = None,
) -> list[pd.DataFrame]:
    """Check skipped entries against all Broadstripes entries for potential duplicates.

    Parameters:
        all_broadstripes: Path to CSV of all Broadstripes entries
        skipped_entries: Path to CSV of skipped entries to check
        all_bs_cols: List of [last, first, middle] column names in all_broadstripes CSV
        skipped_cols: List of [last, first, middle] column names in skipped_entries CSV
        all_bs_demo_cols: Demographic column names in all_broadstripes CSV
            (ordered: Phone, Address, Email, Employer, Degree)
        skipped_demo_cols: Demographic column names in skipped_entries CSV
            (ordered: Phone, Address, Email, Employer, Degree)
        similarity_threshold: Threshold for fuzzy matching (0-1)
        outfile: Optional path to save results CSV

    Returns:
        List of per-match-group DataFrames
    """
    # Load data, validating both name and demographic columns
    existing_df = load_and_validate_csv(
        all_broadstripes,
        all_bs_cols + all_bs_demo_cols,
        "all Broadstripes entries",
    )
    new_df = load_and_validate_csv(
        skipped_entries,
        skipped_cols + skipped_demo_cols,
        "skipped additions",
    )

    # Create standardized names using original format
    names1 = combine_name_parts(new_df, skipped_cols)
    names2 = combine_name_parts(existing_df, all_bs_cols)

    # Find matches
    match_groups = find_potential_matches(names1, names2, similarity_threshold)

    print(f"\nFound {len(match_groups)} potential matches")

    # Build per-group tables
    tables = []
    for group in match_groups:
        table = build_match_table(
            group, new_df, existing_df, skipped_demo_cols, all_bs_demo_cols
        )
        tables.append(table)

    # Print tables
    with pd.option_context(
        "display.max_rows",
        None,
        "display.max_columns",
        None,
        "display.width",
        None,
        "display.max_colwidth",
        40,
    ):
        for table in tables:
            print()
            print(table.to_string(index=False))

    # Save concatenated tables to CSV (blank separator rows between groups)
    if outfile and tables:
        parts = []
        for i, table in enumerate(tables):
            parts.append(table)
            if i < len(tables) - 1:
                blank = pd.DataFrame([[""] * len(table.columns)], columns=table.columns)
                parts.append(blank)
        combined = pd.concat(parts, ignore_index=True)
        combined.to_csv(outfile, index=False)
        print(f"\nWrote results to '{outfile}'")

    return tables


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Check skipped entries against all Broadstripes entries for potential duplicates using fuzzy name matching."
    )
    parser.add_argument(
        "--all_existing",
        type=Path,
        help="Path to CSV file containing all Broadstripes entries",
        required=True,
    )
    parser.add_argument(
        "--skips",
        type=Path,
        help="Path to CSV file containing skipped entries to check",
        required=True,
    )
    parser.add_argument(
        "--all_bs_cols",
        type=str,
        nargs=3,
        help="Names of the last, first, and middle name columns in all_broadstripes CSV (default: 'Last Name First Name Middle Name')",
        default=["Last Name", "First Name", "Middle Name"],
    )
    parser.add_argument(
        "--skipped_cols",
        type=str,
        nargs=3,
        help="Names of the last, first, and middle name columns in skipped_entries CSV (default: 'Last First Middle')",
        default=["Last", "First", "Middle"],
    )
    parser.add_argument(
        "--all_existing_demo_cols",
        type=str,
        nargs=5,
        metavar=("PHONE", "ADDRESS", "EMAIL", "EMPLOYER", "DEGREE"),
        help="Demographic column names in all_broadstripes CSV, in order: phone, address, email, employer, degree",
        default=["Primary Phone", "Address", "Primary Email", "Employer", "Degree"],
    )
    parser.add_argument(
        "--skips_demo_cols",
        type=str,
        nargs=5,
        metavar=("PHONE", "ADDRESS", "EMAIL", "EMPLOYER", "DEGREE"),
        help="Demographic column names in skipped_entries CSV, in order: phone, address, email, employer, degree",
        default=[
            "MP Phone For Address",
            "Address Combined",
            "DND Email",
            "Employer",
            "Degree",
        ],
    )
    parser.add_argument(
        "--similarity_threshold",
        type=float,
        help="Threshold for fuzzy matching (0-1, default: 0.75)",
        default=0.75,
    )
    parser.add_argument(
        "--outfile", type=Path, help="Optional path to save results CSV", default=None
    )

    args = parser.parse_args()

    if args.similarity_threshold < 0 or args.similarity_threshold > 1:
        raise ValueError("Similarity threshold must be between 0 and 1")

    _ = check_skipped_imports(
        all_broadstripes=args.all_existing,
        skipped_entries=args.skips,
        all_bs_cols=args.all_bs_cols,
        skipped_cols=args.skipped_cols,
        all_bs_demo_cols=args.all_existing_demo_cols,
        skipped_demo_cols=args.skips_demo_cols,
        similarity_threshold=args.similarity_threshold,
        outfile=args.outfile,
    )
