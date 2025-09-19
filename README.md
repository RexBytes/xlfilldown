# xlfilldown

[![PyPI version](https://img.shields.io/pypi/v/xlfilldown.svg)](https://pypi.org/project/xlfilldown/)
[![Python versions](https://img.shields.io/pypi/pyversions/xlfilldown.svg)](https://pypi.org/project/xlfilldown/)
[![License](https://img.shields.io/pypi/l/xlfilldown.svg)](https://github.com/RexBytes/xlfilldown/blob/main/LICENSE)


Stream an Excel sheet into **SQLite** or a new **Excel** sheet in constant memory. Forward-fill (pad) selected columns by **header name**, preserve original Excel row numbers, and compute a stable **SHA-256 row hash**.

* Ingests only columns with non-empty headers (from `--header-row`).
* Stores all non-empty values as **TEXT strings** (numbers/dates canonicalized to stable text; strings are stripped; whitespace-only cells become **NULL**).
* Adds optional `excel_row` and `row_hash` columns.
* Streams rows; suitable for large sheets.

---

## Install

### From PyPI (recommended)

```bash
pip install xlfilldown
# or
pipx install xlfilldown
````

Python ≥ 3.9. Depends on `openpyxl`.

## CLI

`xlfilldown` has two subcommands. They share the same **input** options, and differ only in **output** destination.

* `db`   → write to **SQLite**
* `xlsx` → write to **Excel**

### Common input options

* `--infile` *(required)*: Path to input `.xlsx` file.
* `--insheet` *(required)*: Sheet name to read.
* `--header-row` *(required, 1-based)*: Row number containing the headers.

* `--pad-cols`: JSON array of header names to forward-fill.  
  Example: `'["tier1","tier2","tier,4"]'`.

* `--pad-cols-letters`: Alternative to `--pad-cols`.  
  Provide Excel column letters (`A B C AE` etc.). These are resolved to **header names** using `--header-row`.  
  If a referenced column’s header cell is empty (None, whitespace, or “nan”), the command will error.  
  Mutually exclusive with `--pad-cols`.

* `--pad-mode` *(default: `hierarchical`)*: Fill-down strategy.
  * `hierarchical` → **default.** Higher-tier changes reset lower-tier carries.
  * `independent` → legacy/pandas-style ffill. Each padded column carries independently.

* `--drop-blank-rows`: Drop rows where **all** padded columns are empty after padding (treat as spacer rows).
* `--require-non-null`: JSON array of headers; drop the row if **any** are null/blank *after* padding.
* `--row-hash`: Include a `row_hash` column. In DB mode this also creates a non-unique index on `row_hash`.
* `--excel-row-numbers`: Include original Excel row numbers in column `excel_row` (1-based).
* `--if-exists` *(default: `fail`)*: `fail` | `replace` | `append`.

### `db` subcommand (SQLite output)

Additional options:

* `--db` *(required)*: SQLite database file (created if missing).
* `--table`: SQLite table name (default: derived from input sheet name).
* `--batch-size` *(default: 1000)*: Rows per `executemany()` batch.

**Create/append semantics**

* Table columns are: `[row_hash?] [excel_row?] + headers…` (all **`TEXT`**, including `excel_row`).
* If `--if-exists append`, the existing table schema must exactly match the expected column order.
* Helpful indexes are created automatically when enabled: `excel_row` and `row_hash`.

**Examples**

By header names:

```bash
xlfilldown db \
  --infile data.xlsx \
  --insheet "Sheet1" \
  --header-row 1 \
  --pad-cols '["columnname1","columnname2","anothercolumn,3"]' \
  --db out.db
````

By column letters:

```bash
xlfilldown db \
  --infile data.xlsx \
  --insheet "Sheet1" \
  --header-row 1 \
  --pad-cols-letters A C AE \
  --db out.db
```

### `xlsx` subcommand (Excel output)

Additional options:

* `--outfile` *(required)*: Output `.xlsx` file.
* `--outsheet`: Output sheet name (default: derived from input sheet name).

**Sheet-level `--if-exists`**

* `fail`: error if target sheet exists.
* `replace`: recreate target sheet fresh.
* `append`: append below existing rows; the destination header row must match the expected header list (including `excel_row` and/or `row_hash` if enabled).

**Examples**

By header names:

```bash
xlfilldown xlsx \
  --infile data.xlsx \
  --insheet "Sheet1" \
  --header-row 1 \
  --pad-cols '["columnname1","columnname2","anothercolumn,3"]' \
  --outfile out.xlsx \
  --outsheet Processed
```

By column letters:

```bash
xlfilldown xlsx \
  --infile data.xlsx \
  --insheet "Sheet1" \
  --header-row 1 \
  --pad-cols-letters A D \
  --outfile out.xlsx \
  --outsheet Processed
```

---

## Behavior details

### Headers

* Only columns with non-empty header cells on `--header-row` are ingested.
* Empty or duplicate headers after normalization are rejected.

### Forward-fill (padding)

* **Hierarchical (default):**
  Higher-tier changes reset lower-tier carries.
  Example:

  ```
  Tier1   Tier2   Tier3
  apple
         red     sour
  potato
         fried   yellow
  ```

  → produces:

  ```
  apple   red    sour
  potato  None   None
  potato  fried  yellow
  ```

* **Independent (legacy):**
  Each padded column carries independently (pandas-style ffill). Same input produces:

  ```
  apple   red    sour
  potato  red    sour
  potato  fried  yellow
  ```

* Completely empty rows (all headers blank) are preserved as empty **without** applying fill-down; the carry persists past them for later rows.

* Whitespace-only cells are treated as blank.

### Dropping rows

* `--drop-blank-rows`: drops rows where **all** `--pad-cols` are blank (often spacer rows).
* `--require-non-null [A,B,…]`: drops rows where **any** of those headers are blank *after* padding.

### Row hash

* `--row-hash` adds a SHA-256 hex digest over **all ingested columns** (in header order) *after* padding for non-empty rows.
* For completely empty rows, the hash reflects all-empty values (no padding is applied by design).
* SQLite mode creates a non-unique index on `row_hash` for faster lookups.
* Numeric cells are normalized for hashing (e.g., `1`, `1.0` → `1`; no scientific notation).

### Excel row numbers

* `--excel-row-numbers` includes the original Excel row number (1-based) in column `excel_row`.

---

## Python API

```python
from xlfilldown.core import ingest_excel_to_sqlite, ingest_excel_to_excel

# → SQLite
summary = ingest_excel_to_sqlite(
    file="data.xlsx",
    sheet="Sheet1",
    header_row=1,
    pad_cols=["columnname1","columnname2","anothercolumn,3"],
    db="out.db",
    table=None,
    drop_blank_rows=True,
    require_non_null=["columnname1","columnname2"],
    row_hash=True,
    excel_row_numbers=True,
    if_exists="replace",
    batch_size=1000,
    pad_hierarchical=True,   # default
)

# → Excel
summary = ingest_excel_to_excel(
    file="data.xlsx",
    sheet="Sheet1",
    header_row=1,
    pad_cols=["columnname1","columnname2","anothercolumn,3"],
    outfile="out.xlsx",
    outsheet=None,
    drop_blank_rows=True,
    require_non_null=["columnname1","columnname2"],
    row_hash=True,
    excel_row_numbers=True,
    if_exists="replace",
    pad_hierarchical=False,  # use independent fill
)
```

**Return fields**

* SQLite: `{ table, columns, rows_ingested, row_hash, excel_row_numbers }`
* Excel: `{ workbook, sheet, columns, rows_written, row_hash, excel_row_numbers }`

---

## Notes

* All destination columns are written as **`TEXT`** (including `excel_row`). Values are stored as canonical strings; hashing uses the same canonicalization.
* The input workbook is opened with `read_only=True, data_only=True` (formulas are evaluated to cached values).

## License

MIT © RexBytes



