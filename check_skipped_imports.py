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


def find_potential_duplicates(
    names1: list[str], names2: list[str], threshold: float = 0.75
) -> pd.DataFrame:
    """Compare two lists of names to find exact and fuzzy matches.

    Parameters:
        names1: First set of names to compare
        names2: Second set of names to compare
        threshold: Float between 0 and 1 for fuzzy matching similarity threshold

    Returns:
        DataFrame containing potential matches with columns:
        - name: Original name from names1
        - match_type: 'exact' or 'fuzzy'
        - similarity: Float similarity score
        - existing_matches: List of matching names from names2
    """
    results = []
    # Convert inputs to lists of cleaned strings
    names1 = [str(n).lower().strip() for n in names1 if pd.notna(n)]
    names2 = set(str(n).lower().strip() for n in names2 if pd.notna(n))

    for name in names1:
        # Check for exact matches first
        if name in names2:
            results.append(
                {
                    "name": name,
                    "match_type": "exact",
                    "similarity": 1.0,
                    "existing_matches": [name],
                }
            )
            continue
        # If no exact match, check for fuzzy matches
        close_matches = []
        for existing in names2:
            similarity = SequenceMatcher(None, name, existing).ratio()
            if similarity >= threshold:
                close_matches.append((existing, similarity))
        if close_matches:
            close_matches.sort(key=lambda x: x[1], reverse=True)
            results.append(
                {
                    "name": name,
                    "match_type": "fuzzy",
                    "similarity": close_matches[0][1],
                    "existing_matches": [match[0] for match in close_matches],
                }
            )
    return pd.DataFrame(results).sort_values(by="similarity", ascending=False)


def load_and_validate_csv(
    file: Path, name_cols: list[str], file_description: str
) -> pd.DataFrame:
    """Load a CSV file and validate required columns exist.

    Parameters:
        file: Path to CSV file
        name_cols: List of [last, first, middle] column names that must exist
        file_description: Description of file for error messages

    Returns:
        Loaded DataFrame
    """
    df = pd.read_csv(file, dtype=str)

    missing_cols = set(name_cols) - set(df.columns)
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

    Args:
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
            raise ValueError("Invalid name: {parts}")
        combined_names.append(full_name)
    return combined_names


def check_skipped_imports(
    all_broadstripes: Path,
    skipped_entries: Path,
    all_bs_cols: list[str],
    skipped_cols: list[str],
    similarity_threshold: float,
    outfile: Path | None = None,
) -> pd.DataFrame:
    """Check skipped entries against all Broadstripes entries for potential duplicates.

    Parameters:
        all_broadstripes: Path to CSV of all Broadstripes entries
        skipped_entries: Path to CSV of skipped entries to check
        all_bs_cols: List of [last, first, middle] column names in all_broadstripes CSV
        skipped_cols: List of [last, first, middle] column names in skipped_entries CSV
        similarity_threshold: Threshold for fuzzy matching (0-1)
        outfile: Optional path to save results CSV

    Returns:
        DataFrame of potential matches
    """
    # Load data
    existing_df = load_and_validate_csv(
        all_broadstripes, all_bs_cols, "all Broadstripes entries"
    )
    new_df = load_and_validate_csv(skipped_entries, skipped_cols, "skipped additions")

    # Create standardized names using original format
    names1 = combine_name_parts(new_df, skipped_cols)
    names2 = combine_name_parts(existing_df, all_bs_cols)

    # Find matches
    matches = find_potential_duplicates(names1, names2, similarity_threshold)

    print(f"\nFound {len(matches)} potential matches")
    # Print whole df
    with pd.option_context(
        "display.max_rows", None, "display.max_columns", None, "display.width", None
    ):
        print(matches)
    if outfile:
        matches.to_csv(outfile, index=False)
        print(f"\nWrote results to '{outfile}'")
    return matches


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Check skipped entries against all Broadstripes entries for potential duplicates using fuzzy name matching."
    )
    parser.add_argument(
        "--all_broadstripes",
        type=Path,
        help="Path to CSV file containing all Broadstripes entries",
        required=True,
    )
    parser.add_argument(
        "--skipped_entries",
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
        all_broadstripes=args.all_broadstripes,
        skipped_entries=args.skipped_entries,
        all_bs_cols=args.all_bs_cols,
        skipped_cols=args.skipped_cols,
        similarity_threshold=args.similarity_threshold,
        outfile=args.outfile,
    )
