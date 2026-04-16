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

## Known-failing tests

A couple of tests in `test_parse_fullnames.py` are marked
`@pytest.mark.xfail(strict=True)`. They codify intended behavior for a known
bug in `parse_fullnames` (tokens ending in `.` are always treated as middle
initial, which misclassifies suffixes like `"Jr."` and `"Sr."`). Once the bug
is fixed, those tests will flip from xfail to pass and pytest will fail
loudly on the now-unnecessary marker — that's the signal to remove it.
