# Real BU-list fixtures

This directory is **gitignored**. You can drop real employer `.xlsx` files here
and the tests in `test_real_data.py` will run against them automatically. If the
directory is empty, those tests are skipped.

## What to put here

1. One or more `.xlsx` files exported from the employer. Each filename **must
   contain the received date in `YYYY.MM.DD` format** — this is the same
   requirement the production script enforces (`get_list_identifier`).

2. A `program_mapping.toml` file. The simplest option is to symlink or copy the
   production `program_mapping.toml` from the repo root:

   ```bash
   ln -s ../../../../program_mapping.toml program_mapping.toml
   ```

3. (Optional) For each xlsx, you can provide `<xlsx-stem>.params.toml` alongside
   it to override the per-file arguments to `parse_dartmouth_bu`. Without it,
   the tests use sensible defaults (see `test_real_data.py`). Example:

   ```toml
   program_col = "Program/Field of Study"
   fullname_col = "FULL_NAME (LFM)"
   # Optional — defaults to ("ADDRESS_LINE1", "ADDRESS_LINE2", "TOWN/CITY", "ST", "ZIP"):
   # address_cols = ["ADDRESS_LINE1", "ADDRESS_LINE2", "TOWN/CITY", "ST", "ZIP"]
   # Optional — use this instead of fullname_col for older LFM-split lists:
   # lfm_cols = ["LAST", "FIRST", "MIDDLE"]
   ```
