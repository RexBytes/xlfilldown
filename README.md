# xlfilldown

[![PyPI version](https://img.shields.io/pypi/v/xlfilldown.svg)](https://pypi.org/project/xlfilldown/)
[![Python versions](https://img.shields.io/pypi/pyversions/xlfilldown.svg)](https://pypi.org/project/xlfilldown/)
[![License](https://img.shields.io/pypi/l/xlfilldown.svg)](https://github.com/RexBytes/xlfilldown/blob/main/LICENSE)
[![Documentation](https://img.shields.io/badge/docs-guide-blue.svg)](https://rexbytes.com/2025/09/23/xlfilldown-the-complete-guide/)
[![Downloads](https://static.pepy.tech/badge/xlfilldown)](https://pepy.tech/project/xlfilldown)
[![Socket Badge](https://badge.socket.dev/pypi/package/xlfilldown/1.0.3?artifact_id=tar-gz)](https://badge.socket.dev/pypi/package/xlfilldown/1.0.3?artifact_id=tar-gz)
[![Changelog](https://img.shields.io/badge/changelog-releases-orange.svg)](https://github.com/RexBytes/xlfilldown/releases)


Stream an Excel sheet into **SQLite** or a new **Excel** sheet in constant memory. Forward-fill selected columns by **header name**, preserve original Excel row numbers, and compute a stable **SHA-256 row hash**.

* Ingests only columns with non-empty headers (from `--header-row`).
* Stores all non-empty values as **TEXT strings** (numbers/dates canonicalized to stable text; strings are stripped; whitespace-only cells become **NULL**).
* Optional `excel_row` and `row_hash` columns.
* Streams rows; suitable for large sheets.

---

## Install

### From PyPI (recommended)

```bash
pip install xlfilldown
# or
pipx install xlfilldown
```

Python ≥ 3.9. Depends on `openpyxl`.

---

## CLI

`xlfilldown` has two subcommands that share the same **input** options and differ only in the **output** destination:

* `db`   → write to **SQLite**
* `xlsx` → write to **Excel**

### Common input options

* `--infile` *(required)*: Path to input `.xlsx` file.

* `--insheet` *(required)*: Sheet name to read.

* `--header-row` *(required, 1-based)*: Row number containing the headers.

* `--fill-cols`: JSON array of header names to forward-fill.
  Example: `'["columnname1","columnname2","anothercolumn,3"]'`.

* `--fill-cols-letters`: Alternative to `--fill-cols`.
  Provide Excel column letters (`A B C AE` etc.). These are resolved to **header names** using `--header-row`.
  If a referenced column’s header cell is empty (None, whitespace, or “nan”), the command errors.
  Mutually exclusive with `--fill-cols`.

* `--fill-mode` *(default: `hierarchical`)*: Fill strategy.

  * `hierarchical` → Higher-tier column changes **reset** lower-tier carries.
  * `independent` → Pandas-style `ffill`, each listed column carries independently.

* `--drop-blank-rows`: Drop rows where **all** fill columns are empty *after* filling (treat as spacer rows).

* `--require-non-null`: JSON array of headers; drop the row if **any** are null/blank *after* fill.

* `--require-non-null-letters`: Excel column letters; resolved to headers and **merged** with `--require-non-null`.

* `--row-hash`: Include a `row_hash` column. In DB mode this also creates a non-unique index on `row_hash`.

* `--excel-row-numbers`: Include original Excel row numbers in column `excel_row` (1-based).

* `--if-exists` *(default: `fail`)*: `fail` | `replace` | `append`.

> **Header matching:** After normalization (trim; case preserved; `'nan'` → blank), names must match exactly.

---

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
  --fill-cols '["columnname1","columnname2","anothercolumn,3"]' \
  --db out.db
```

By column letters:

```bash
xlfilldown db \
  --infile data.xlsx \
  --insheet "Sheet1" \
  --header-row 1 \
  --fill-cols-letters A C AE \
  --db out.db
```

---

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
  --fill-cols '["columnname1","columnname2","anothercolumn,3"]' \
  --outfile out.xlsx \
  --outsheet Processed
```

By column letters:

```bash
xlfilldown xlsx \
  --infile data.xlsx \
  --insheet "Sheet1" \
  --header-row 1 \
  --fill-cols-letters A D \
  --outfile out.xlsx \
  --outsheet Processed
```

---

## Behavior details

### Headers

* Only columns with non-empty header cells on `--header-row` are ingested.
* Empty or duplicate headers after normalization are rejected.

### Forward-fill (filling)

* **Hierarchical (default): order matters**
  The hierarchy is the **order you pass the columns**. The leftmost is the highest tier.

  * Names: `--fill-cols '["Region","Country","City"]'` ⇒ Region > Country > City
  * Letters: `--fill-cols-letters A C B` ⇒ Column A > Column C > Column B
    When a higher-tier value appears on a row, all lower-tier carries **reset for that row**.

* **Independent (pandas-style `ffill`)**
  Each listed column forward-fills **independently**.
  Order of columns **does not** matter. Columns do **not** reset each other.

* Completely empty rows (all headers blank) are preserved as empty **without** applying fill; the carry persists past them for later rows.

* Whitespace-only cells are treated as blank.

**Illustration**

Input:

```
columnname1   columnname2   anothercolumn,3
apple
       red     sour
potato
       fried   yellow
```

Hierarchical output:

```
apple   red    sour
potato  None   None
potato  fried  yellow
```

Independent output:

```
apple   red    sour
potato  red    sour
potato  fried  yellow
```

### Dropping rows

* `--drop-blank-rows`: drops rows where **all** `--fill-cols` are blank (often spacer rows).
* `--require-non-null [A,B,…]` / `--require-non-null-letters`: drops rows where **any** of those headers are blank *after* filling.

### Row hash

* `--row-hash` adds a SHA-256 hex digest over **all ingested columns** (in header order) *after* filling for non-empty rows.
* For completely empty rows, the hash reflects all-empty values (no filling is applied by design).
* SQLite mode creates a non-unique index on `row_hash` for faster lookups.
* Numeric cells are normalized for hashing (e.g., `1`, `1.0` → `1`; no scientific notation).

### Excel row numbers

* `--excel-row-numbers` includes the original Excel row number (1-based) in column `excel_row`.

---

## Python API

```python
from xlfilldown.api import ingest_excel_to_sqlite, ingest_excel_to_excel

# → SQLite
summary = ingest_excel_to_sqlite(
    file="data.xlsx",
    sheet="Sheet1",
    header_row=1,
    # choose one:
    fill_cols=["columnname1", "columnname2", "anothercolumn,3"],   # by header names
    # fill_cols_letters=["A", "B", "C"],                           # or by Excel letters
    db="out.db",
    table=None,
    drop_blank_rows=True,
    # choose one (or both, merged & de-duped):
    require_non_null=["columnname1", "columnname2"],               # by header names
    # require_non_null_letters=["A", "B"],                         # or by Excel letters
    row_hash=True,
    excel_row_numbers=True,
    if_exists="replace",
    batch_size=1000,
    fill_mode="hierarchical",    # default hierarchical fill
    # fill_mode="independent",   # independent (pandas-style) fill
)

# → Excel
summary = ingest_excel_to_excel(
    file="data.xlsx",
    sheet="Sheet1",
    header_row=1,
    fill_cols=["columnname1", "columnname2", "anothercolumn,3"],
    # or: fill_cols_letters=["A", "B", "C"],
    outfile="out.xlsx",
    outsheet=None,
    drop_blank_rows=True,
    require_non_null=["columnname1", "columnname2"],
    # or: require_non_null_letters=["A", "B"],
    row_hash=True,
    excel_row_numbers=True,
    if_exists="replace",
    fill_mode="independent",     # independent (pandas-style) fill
    # fill_mode="hierarchical",  # hierarchical (default)
)

````

**Return fields**

* SQLite: `{ table, columns, rows_ingested, row_hash, excel_row_numbers }`
* Excel: `{ workbook, sheet, columns, rows_written, row_hash, excel_row_numbers }`


---
## Raw Ingest Examples

Sometimes you may want to ingest a sheet *exactly as-is*, without applying any fill-down logic.  
This is useful if you just want to load the data into SQLite or Excel with audit columns (`row_hash`, `excel_row`) and handle nulls/blanks later.

---

### CLI: Raw ingest to SQLite
```bash
xlfilldown db \
  --infile data.xlsx --insheet Sheet1 --header-row 1 \
  --db out.db --table raw_ingest \
  --ingest-mode raw \
  --row-hash --excel-row-numbers --if-exists replace
````

---

### CLI: Raw ingest to Excel

```bash
xlfilldown xlsx \
  --infile data.xlsx --insheet Sheet1 --header-row 1 \
  --outfile out.xlsx --outsheet RawSheet \
  --ingest-mode raw \
  --row-hash --excel-row-numbers --if-exists replace
```

---

### Python API: Raw ingest to SQLite

```python
from xlfilldown.api import ingest_excel_to_sqlite

summary = ingest_excel_to_sqlite(
    file="data.xlsx", sheet="Sheet1", header_row=1,
    ingest_mode="raw",                      # skip fill-down
    db="out.db", table="raw_ingest",
    row_hash=True, excel_row_numbers=True,
    if_exists="replace",
)
print(summary)
```

---

### Python API: Raw ingest to Excel

```python
from xlfilldown.api import ingest_excel_to_excel

summary = ingest_excel_to_excel(
    file="data.xlsx", sheet="Sheet1", header_row=1,
    ingest_mode="raw",                      # skip fill-down
    outfile="out.xlsx", outsheet="RawSheet",
    row_hash=True, excel_row_numbers=True,
    if_exists="replace",
)
print(summary)
```



## Notes

* All destination columns are written as **`TEXT`** (including `excel_row`). Values are stored as canonical strings; hashing uses the same canonicalization.
* The input workbook is opened with `read_only=True, data_only=True` (formulas use cached values).

## License

MIT © RexBytes

## Proof of Ownership

Woof woof!

RexBytes here — author, creator, and maintainer of this project.  
I don't know any secret handshakes, so here’s a signed dog tag instead.

```
-----BEGIN PGP SIGNED MESSAGE-----
Hash: SHA512

woof woof!

RexBytes here - author, creator, and maintainer of project xlfilldown.

In case we ever meet at the same park, I'll show you my signed dog tag to say hello.

Signed on: 2025-10-01

-----BEGIN PGP SIGNATURE-----

iHUEARYKAB0WIQT8gR2B/yYm6t41oRy0jpEjfRE3bwUCaNwumAAKCRC0jpEjfRE3
b9hPAP43wgEP/kKM5cAEFxzSvsFEEjYtVWCTDLAZMe4IDEfzdQD/QXHbQg2WRTMP
Nq2fPFhHzdTJd7/3HjI/GyLWVR5kmQ0=
=KYYE
-----END PGP SIGNATURE-----
```

## Donations

This little package started as a way to scratch my own itch.
Turns out, it saved me from countless Excel headaches 🧹📊.  

If xlfilldown also makes your life easier, and you feel like tossing a treat into the dog bowl,  
you can send a few sats my way 🐶🍺.  

**BTC (Bech32):** `bc1qk2d5cnn5xlnerxq3fcfrsnutzqmye0fe020rdd`

Totally optional. I’ll only use it for beer, toys, maybe even more coding.  
Either way, I hope xlfilldown helps flatten your spreadsheets and your stress. 🐾


