"""Check the "skips" from a Broadstripes data import for workers we might already have.

When we import a parsed BU list, Broadstripes matches workers by email and
"skips" the rows it can't match. Most skips are new workers, but some are
existing workers whose email changed (often along with their name). This script
compares each skipped row against a full export of Broadstripes and lists
possible matches, by name (exact or fuzzy) and by phone number.

Run `python check_skipped_imports.py --help` for usage, and see README.md for a
worked example. The overall listwork process this fits into is documented in
process_employer_list_doc.md.
"""

try:
    import pandas as pd
except ImportError as e:
    raise ImportError(
        "This script requires the python library pandas. Please install it to your environment."
    ) from e

import argparse
import re
import sys
from difflib import SequenceMatcher
from pathlib import Path

pd.options.mode.copy_on_write = True

# Defaults match the output of parse_employer_bu.py (for the skips file) and
# the "Contact and Degree Info" layout export (for the Broadstripes file).
DEFAULT_SKIPS_NAME_COLS = ("Last", "First", "Middle")
DEFAULT_BS_NAME_COLS = ("Last Name", "First Name", "Middle Name")
DEFAULT_SKIPS_INFO_COLS = ("MP Phone For Address", "DND Email", "Employer")
DEFAULT_BS_INFO_COLS = ("Primary Phone", "Primary Email", "Employer")
DEFAULT_BS_ID_COL = "Broadstripes ID"

# The info columns are shown side by side under these headings. Phone is also
# used for matching.
INFO_FIELDS = ("Phone", "Email", "Employer")

# Match reasons. "Strong" reasons are very likely to be the same person.
SAME_NAME = "same name"
SAME_PHONE = "same phone"
SIMILAR_NAME = "similar name"
STRONG_REASONS = (SAME_NAME, SAME_PHONE)


def check_columns(df: pd.DataFrame, wanted: dict[str, str], file_label: str) -> None:
    """Check up front that every column we were asked to use actually exists.

    `wanted` maps an argument name (e.g. "--skips_name_cols") to a column name.
    All missing columns are reported at once, alongside the file's real columns.
    """
    missing = {arg: col for arg, col in wanted.items() if col not in df.columns}
    if not missing:
        return
    print(f"\nError: some requested column(s) are not in the {file_label} file:")
    for arg, col in missing.items():
        print(f"  {arg}: '{col}'")
    print(f"The {file_label} file's actual columns are:")
    for col in df.columns:
        print(f"  '{col}'")
    print("Fix the argument(s) above to match (see --help).")
    raise ValueError(
        f"Column(s) not found in {file_label} file: {sorted(missing.values())}"
    )


def load_csv(file: Path, file_label: str) -> pd.DataFrame:
    df = pd.read_csv(file, dtype=str)
    print(f"Loaded {file_label} file '{file}' ({df.shape[0]} rows).")
    return df


def format_name(last, first, middle) -> str:
    """Combine name parts as "Last, First Middle". Returns "" if there's no name."""
    last, first, middle = (
        str(x).strip() if pd.notna(x) else "" for x in (last, first, middle)
    )
    if not (last or first):
        return ""
    return f"{last}, {first} {middle}".strip()


def normalize_phone(phone) -> str:
    """Reduce a phone number to its digits, dropping a leading US country code.

    Returns "" for missing or implausibly short numbers, so they never match.
    """
    if pd.isna(phone):
        return ""
    digits = re.sub(r"\D", "", str(phone))
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits if len(digits) >= 7 else ""


def find_potential_matches(
    skip_names: list[str],
    bs_names: list[str],
    skip_phones: list[str],
    bs_phones: list[str],
    threshold: float = 0.75,
) -> list[dict]:
    """Find Broadstripes records that might be the same person as each skipped row.

    Names are compared case-insensitively. A Broadstripes record is a possible
    match if it has the same name, a similar name (similarity >= threshold), or
    the same phone number. Phones should already be normalized.

    Returns one group per skipped row that has any possible matches:
        - skip_idx: position in skip_names
        - strong: whether any match is a same name or same phone match
        - matches: list of dicts (bs_idx, reasons, similarity), strongest first
    Groups are sorted strongest first, then by best name similarity.
    """
    bs_lower = [n.lower() for n in bs_names]
    bs_by_phone: dict[str, list[int]] = {}
    for idx, phone in enumerate(bs_phones):
        if phone:
            bs_by_phone.setdefault(phone, []).append(idx)

    groups = []
    for skip_idx, (name, phone) in enumerate(zip(skip_names, skip_phones)):
        name = name.lower()
        found: dict[int, dict] = {}

        if name:
            # SequenceMatcher caches info about seq2, so keep the skip name there
            matcher = SequenceMatcher(autojunk=False)
            matcher.set_seq2(name)
            for bs_idx, bs_name in enumerate(bs_lower):
                if not bs_name:
                    continue
                if bs_name == name:
                    found[bs_idx] = {"reasons": [SAME_NAME], "similarity": 1.0}
                    continue
                matcher.set_seq1(bs_name)
                # Cheap upper bounds first; ratio() is the slow part
                if (
                    matcher.real_quick_ratio() >= threshold
                    and matcher.quick_ratio() >= threshold
                ):
                    similarity = matcher.ratio()
                    if similarity >= threshold:
                        found[bs_idx] = {
                            "reasons": [SIMILAR_NAME],
                            "similarity": similarity,
                        }

        for bs_idx in bs_by_phone.get(phone, []) if phone else []:
            if bs_idx in found:
                found[bs_idx]["reasons"].insert(0, SAME_PHONE)
            else:
                similarity = SequenceMatcher(
                    None, name, bs_lower[bs_idx], autojunk=False
                ).ratio()
                found[bs_idx] = {"reasons": [SAME_PHONE], "similarity": similarity}

        if not found:
            continue
        matches = [{"bs_idx": idx, **m} for idx, m in found.items()]
        for m in matches:
            m["strong"] = any(r in STRONG_REASONS for r in m["reasons"])
        matches.sort(key=lambda m: (m["strong"], m["similarity"]), reverse=True)
        groups.append(
            {
                "skip_idx": skip_idx,
                "strong": matches[0]["strong"],
                "matches": matches,
            }
        )

    groups.sort(
        key=lambda g: (g["strong"], g["matches"][0]["similarity"]), reverse=True
    )
    return groups


def build_match_table(
    group: dict,
    skip_names: list[str],
    bs_names: list[str],
    skip_df: pd.DataFrame,
    bs_df: pd.DataFrame,
    skip_info_cols: list[str],
    bs_info_cols: list[str],
    bs_id_col: str | None,
) -> pd.DataFrame:
    """Build a small table: one skipped row, followed by its possible matches."""

    def info(df, idx, cols):
        values = df.iloc[idx][list(cols)]
        return {f: (v if pd.notna(v) else "") for f, v in zip(INFO_FIELDS, values)}

    skip_idx = group["skip_idx"]
    rows = [
        {
            "": f"Skip (CSV row {skip_idx + 2})",
            "Name": skip_names[skip_idx],
            **info(skip_df, skip_idx, skip_info_cols),
            "Broadstripes ID": "",
            "Name similarity": "",
        }
    ]
    for m in group["matches"]:
        bs_idx = m["bs_idx"]
        bs_id = bs_df.iloc[bs_idx][bs_id_col] if bs_id_col else ""
        rows.append(
            {
                "": "Match: " + " + ".join(m["reasons"]),
                "Name": bs_names[bs_idx],
                **info(bs_df, bs_idx, bs_info_cols),
                "Broadstripes ID": bs_id if pd.notna(bs_id) else "",
                "Name similarity": f"{m['similarity']:.2f}",
            }
        )
    return pd.DataFrame(rows)


def check_skipped_imports(
    skips_file: Path,
    broadstripes_file: Path,
    skips_name_cols: tuple[str, ...] | list[str] = DEFAULT_SKIPS_NAME_COLS,
    bs_name_cols: tuple[str, ...] | list[str] = DEFAULT_BS_NAME_COLS,
    skips_info_cols: tuple[str, ...] | list[str] = DEFAULT_SKIPS_INFO_COLS,
    bs_info_cols: tuple[str, ...] | list[str] = DEFAULT_BS_INFO_COLS,
    threshold: float = 0.75,
    outfile: Path | None = None,
    write: bool = True,
) -> list[pd.DataFrame]:
    """Compare skipped rows against all Broadstripes records and report possible matches.

    Prints each skipped row that has possible matches, and (if `write`) saves
    them all to a CSV. By default the CSV is written next to the skips file.
    Returns the per-skip tables.
    """
    skip_df = load_csv(skips_file, "skips")
    bs_df = load_csv(broadstripes_file, "Broadstripes")

    def wanted(arg_prefix, name_cols, info_cols) -> dict[str, str]:
        # e.g. {"--skips_name_cols (last)": "Last", ...}
        name_args = [
            f"--{arg_prefix}_name_cols ({p})" for p in ("last", "first", "middle")
        ]
        info_args = [f"--{arg_prefix}_info_cols ({f.lower()})" for f in INFO_FIELDS]
        return dict(zip(name_args + info_args, [*name_cols, *info_cols]))

    check_columns(skip_df, wanted("skips", skips_name_cols, skips_info_cols), "skips")
    check_columns(
        bs_df, wanted("broadstripes", bs_name_cols, bs_info_cols), "Broadstripes"
    )
    # The ID is only shown for convenience, so don't insist on it
    bs_id_col = DEFAULT_BS_ID_COL if DEFAULT_BS_ID_COL in bs_df.columns else None

    def names(df, cols) -> list[str]:
        return [format_name(*parts) for parts in zip(*(df[c] for c in cols))]

    skip_names = names(skip_df, skips_name_cols)
    bs_names = names(bs_df, bs_name_cols)
    skip_phones = [normalize_phone(p) for p in skip_df[skips_info_cols[0]]]
    bs_phones = [normalize_phone(p) for p in bs_df[bs_info_cols[0]]]

    groups = find_potential_matches(
        skip_names, bs_names, skip_phones, bs_phones, threshold
    )
    tables = [
        build_match_table(
            g,
            skip_names,
            bs_names,
            skip_df,
            bs_df,
            skips_info_cols,
            bs_info_cols,
            bs_id_col,
        )
        for g in groups
    ]

    n_strong = sum(g["strong"] for g in groups)
    print(
        f"\nCompared {len(skip_df)} skipped rows against {len(bs_df)} Broadstripes records."
        f"\n  {len(skip_df) - len(groups)} skips have no possible match (likely new workers)."
        f"\n  {len(groups)} skips have possible matches:"
        f"\n    {n_strong} strong (same name or same phone number). Check these carefully."
        f"\n    {len(groups) - n_strong} weaker (similar name only). Most of these are different people."
    )
    if not groups:
        return tables

    with pd.option_context(
        "display.max_columns", None, "display.width", None, "display.max_colwidth", 45
    ):
        for i, (group, table) in enumerate(zip(groups, tables)):
            if i == 0 and group["strong"]:
                print("\n=== Strong possible matches ===")
            if not group["strong"] and (i == 0 or groups[i - 1]["strong"]):
                print("\n=== Weaker possible matches (similar name only) ===")
            print()
            print(table.to_string(index=False))

    if write:
        if not outfile:
            outfile = skips_file.parent / f"{skips_file.stem} - possible duplicates.csv"
        # Blank row between groups, to make the CSV easier to read
        blank = pd.DataFrame([[""] * len(tables[0].columns)], columns=tables[0].columns)
        parts = [part for table in tables for part in (table, blank)][:-1]
        outfile.parent.mkdir(parents=True, exist_ok=True)
        pd.concat(parts, ignore_index=True).to_csv(outfile, index=False)
        print(f"\nWrote these possible matches to '{outfile}'")
    return tables


def _threshold(value: str) -> float:
    x = float(value)
    if not 0 <= x <= 1:
        raise argparse.ArgumentTypeError("must be between 0 and 1")
    return x


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Check the skips from a Broadstripes data import against a full export of "
            "Broadstripes, to find workers who might already be in our database "
            "(e.g. because their name or email changed). Possible matches are found "
            "by same name, similar name, or same phone number."
        ),
    )
    parser.add_argument(
        "-s",
        "--skips",
        type=Path,
        required=True,
        help="Path to the skips CSV downloaded from the Broadstripes data import.",
    )
    parser.add_argument(
        "-b",
        "--broadstripes",
        type=Path,
        required=True,
        help="Path to the CSV export of everyone in Broadstripes (see process_employer_list_doc.md).",
    )
    parser.add_argument(
        "-o",
        "--outfile",
        type=Path,
        default=None,
        help="Path to write the possible matches (.csv). Optional. By default, written next to the skips file.",
    )
    parser.add_argument(
        "--threshold",
        type=_threshold,
        default=0.75,
        help="How similar two names must be (0-1) to count as a possible match. Lower finds more, but with more false alarms. Default: 0.75",
    )
    parser.add_argument(
        "--skips_name_cols",
        nargs=3,
        metavar=("LAST", "FIRST", "MIDDLE"),
        default=list(DEFAULT_SKIPS_NAME_COLS),
        help="Name columns in the skips file. Default: %(default)s (as made by parse_employer_bu.py)",
    )
    parser.add_argument(
        "--broadstripes_name_cols",
        nargs=3,
        metavar=("LAST", "FIRST", "MIDDLE"),
        default=list(DEFAULT_BS_NAME_COLS),
        help="Name columns in the Broadstripes file. Default: %(default)s",
    )
    parser.add_argument(
        "--skips_info_cols",
        nargs=3,
        metavar=("PHONE", "EMAIL", "EMPLOYER"),
        default=list(DEFAULT_SKIPS_INFO_COLS),
        help="Phone, email, and employer columns in the skips file. Phone is used for matching; all three are shown in the output. Default: %(default)s",
    )
    parser.add_argument(
        "--broadstripes_info_cols",
        nargs=3,
        metavar=("PHONE", "EMAIL", "EMPLOYER"),
        default=list(DEFAULT_BS_INFO_COLS),
        help="Phone, email, and employer columns in the Broadstripes file. Default: %(default)s",
    )
    return parser


def main(argv: list[str] | None = None) -> None:
    args = build_parser().parse_args(argv)
    check_skipped_imports(
        skips_file=args.skips,
        broadstripes_file=args.broadstripes,
        skips_name_cols=args.skips_name_cols,
        bs_name_cols=args.broadstripes_name_cols,
        skips_info_cols=args.skips_info_cols,
        bs_info_cols=args.broadstripes_info_cols,
        threshold=args.threshold,
        outfile=args.outfile,
    )


if __name__ == "__main__":
    main(sys.argv[1:])
