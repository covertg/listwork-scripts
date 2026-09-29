# parse_employer_bu tests

Pytest suite for `parse_employer_bu.py`. Run from the repo root:

```bash
pytest tests/parse_employer_bu/
```

Install test dependencies if you don't already have them:

```bash
pip install -r requirements-dev.txt
```

## Fixtures

- `fixtures/sample_list.xlsx` — small synthetic employer list

- `fixtures/sample_program_mapping.toml` — test mapping that covers every
  program used by `sample_list.xlsx`.

- `fixtures/real/` — **gitignored**. Drop real employer `.xlsx` files here
  along with a `program_mapping.toml` to exercise the parser against real
  data. See `fixtures/real/README.md`.

## Test modules

| Module | Covers |
| --- | --- |
| `test_load_list.py` | reading the xlsx, whitespace stripping, zip leading zeros, multi-sheet warning |
| `test_extract_date.py` / `test_get_list_identifier.py` | pulling `YYYY.MM.DD` out of the filename |
| `test_convert_program_mapping.py` | the `program_mapping.toml` lookup and the unknown-code error |
| `test_parse_fullnames.py` | splitting `"Last, First M."` into three columns |
| `test_str_combine.py` / `test_make_address_combined.py` | the combined address column |
| `test_check_columns.py` | the fail-fast guard on mistyped `--*_col` arguments |
| `test_check_data_quality.py` | the advisory warnings for duplicate / shifted / malformed rows |
| `test_cli.py` | argument validation, `--show_columns`, default output path |
| `test_integration.py` | the whole pipeline against `sample_list.xlsx` |
| `test_real_data.py` | the whole pipeline against real lists in `fixtures/real/` (skipped if absent) |

## No known-failing tests

There used to be two `xfail(strict=True)` tests here codifying a suffix-parsing
bug in `parse_fullnames`; the bug is fixed and they now pass normally. If you
find yourself documenting a bug this way again, strict xfail is a good pattern:
the test flips to a loud failure the moment someone fixes the bug, which is the
signal to drop the marker.
