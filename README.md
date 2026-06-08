# tariffs_excel2sql_tools

Small command-line tools that turn an **Excel worksheet of tariff rows** into the
**SQL `VALUES` rows** used to populate the GTM/ONESOURCE Global Trade
`tmdHTSAdditional` table.

They exist to remove the tedious, error-prone part of a "DBA insert" data-maintenance
story — hand-typing hundreds of `(N'...', N'...', ...)` tuples — by generating them
directly from the approved spreadsheet that ships with the work item.

| File | Language | Role |
|------|----------|------|
| [main.py](main.py) | Python 3 + pandas | **Generator** — reads an Excel sheet, emits one SQL `VALUES` tuple per row |
| [tool.js](tool.js) | Node.js (no deps) | **Post-processor** — `batch` / `index` / `merge` the generated rows |
| [.command](.command) | text | A reminder of the `main.py` invocation syntax |

> ⚠️ These tools generate the **bulk row data only**. They do **not** produce the
> final, production-grade, idempotent deployment script, and they do **not** produce
> the verification script. See [What these tools do *not* do](#what-these-tools-do-not-do).

---

## Table of contents

- [Background — where this fits](#background--where-this-fits)
  - [The `tmdHTSAdditional` column contract](#the-tmdhtsadditional-column-contract)
- [`main.py` — Excel → SQL `VALUES` rows](#mainpy--excel--sql-values-rows)
  - [Purpose](#purpose)
  - [Requirements](#requirements)
  - [Arguments](#arguments)
  - [Usage](#usage)
  - [How it works](#how-it-works)
  - [Output](#output)
- [`tool.js` — post-process the generated rows](#tooljs--post-process-the-generated-rows)
  - [`batch` — split rows into runnable `INSERT` batches](#batch--split-rows-into-runnable-insert-batches)
  - [`index` — add a `RowId` to each tuple](#index--add-a-rowid-to-each-tuple)
  - [`merge` — concatenate files in a folder](#merge--concatenate-files-in-a-folder)
- [End-to-end workflow](#end-to-end-workflow)
- [What these tools do *not* do](#what-these-tools-do-not-do)
- [Caveats / known rough edges](#caveats--known-rough-edges)
- [Quick reference](#quick-reference)

---

## Background — where this fits

This toolset supports stories such as
[US 5441030 — "Regulatory- DBA Insert tmdhtsAdditional - Taiwan Section 232 New Chapter 99 Headings"](https://dev.azure.com/tr-tax/TaxTrade/_workitems/edit/5441030).

The pattern for these stories is:

1. A regulatory change (e.g. a Federal Register Notice / CSMS message) requires a set
   of HTS codes to be added to the `tmdHTSAdditional` table for a given country,
   tariff type, and Chapter 99 heading.
2. The approved record set is attached to the work item as an **Excel file** (e.g.
   `tmdhtsAdditional_Taiwan_232_Changes_5.xlsx`), with an **`InsertData`** tab that
   holds every column for every row (386 rows in US 5441030).
3. A DBA deployment SQL script is produced (committed to the
   [`gtm-legacy_gtm-sql`](https://github.com/tr/gtm-legacy_gtm-sql) repo, e.g.
   [`V26.2.0710__DATA_tmdHTSAdditional_Taiwan_232_Chapter99.sql`](https://github.com/tr/gtm-legacy_gtm-sql/blob/e354f92265d5bf7dd23de454e1c3a5403da01ef2/Database/Application/26.2/Hotfix/V26.2.0710__DATA_tmdHTSAdditional_Taiwan_232_Chapter99.sql))
   and a read-only **verification script** that checks the work item's acceptance
   criteria.

**`main.py` automates step 2 → the row data needed by step 3.** It takes the
`InsertData` sheet and produces the `(...)`-style `VALUES` rows. A developer/DBA then
embeds those rows into the production template (backup, idempotent insert, `NOLOCK`,
transaction, error handling — none of which these tools generate).

### The `tmdHTSAdditional` column contract

Both tools are hard-wired to these ten columns, in this order:

```
HTSNum, Chapter99, CountryofOrigin, StartEffDate, EndEffDate,
TariffType, TariffGroup, RequiredStatusCode, ValidationLevel, ExportDate
```

The Excel sheet **must** contain columns with exactly these names (leading/trailing
whitespace is stripped). Any extra columns are ignored; if any of the ten are missing
the script fails.

---

## `main.py` — Excel → SQL `VALUES` rows

### Purpose

Read one worksheet of an Excel workbook and write a `.sql` file where each data row
becomes a single SQL `VALUES` tuple, prefixed with a `SELECT … INTO <table> FROM (VALUES`
header line.

### Requirements

```bash
pip install pandas openpyxl    # openpyxl is pandas' engine for .xlsx
```

Python 3.

### Arguments

| Flag | Required | Default | Meaning |
|------|----------|---------|---------|
| `--excel` | ✅ | — | Path to the Excel workbook (`.xlsx`) |
| `--sheet` | ✅ | — | Worksheet/tab name to read (e.g. `InsertData`) |
| `--table` | ❌ | `[dbo].[tmdhtsAdditional]` | Target table name used in the header line |
| `--out`   | ✅ | — | Path of the `.sql` file to write |

### Usage

```bash
python main.py --excel tmdhtsAdditional_Taiwan_232_Changes_5.xlsx \
               --sheet InsertData \
               --table "[dbo].[tmdHTSAdditional]" \
               --out ROWS.sql
```

(The same invocation is recorded in [.command](.command).)

On success it prints, e.g.:

```
Wrote 386 statements to ROWS.sql
```

### How it works

1. **Load** the sheet with `pandas.read_excel(..., dtype=object)` so every cell keeps
   its raw type.
2. **Normalize** column headers (strips surrounding whitespace).
3. **Select** exactly the ten contract columns (`df[cols]`) and **drop duplicate rows**.
4. **Emit a header** line via `get_insert_command()`:
   `SELECT [RowId], [HTSNum], … , [ExportDate] INTO <table> FROM (VALUES`
5. **Emit one tuple per row** via `make_insert()` → `,(val1, val2, …, val10)`.
   - The **first** data row has its leading comma stripped.
   - The **last** data row gets a closing `);` appended.
6. **Write** the joined text to `--out` (UTF-8).

#### Value formatting (`sql_literal` / `clean_number`)

Each cell is converted to a SQL literal according to its type:

| Cell value | Emitted as | Notes |
|------------|-----------|-------|
| empty / `NaN` / blank string | `N''` | **emits an empty string, *not* `NULL`** |
| `pandas.Timestamp` | `CAST(N'YYYY-MM-DD HH:MM:SS' AS DATETIME)` | time component preserved |
| integer-valued number (`850132.0`) | `N'850132'` | trailing `.0` stripped → keeps HTS codes dot-free |
| non-integer number | `N'<value>'` | |
| string that parses as a date | `CAST(N'…' AS DATETIME)` | tries `%Y-%m-%d %H:%M:%S`, `%Y-%m-%d`, `%m/%d/%Y %H:%M`, `%m/%d/%Y` |
| any other string | `N'<escaped>'` | single quotes doubled (`'` → `''`) |

The float→int normalization is the key feature for HTS numbers: Excel commonly reads a
code like `850132` as the float `850132.0`, which would otherwise produce a value
containing a dot. `clean_number()` strips it, satisfying the "HTSNum must contain digits
only" rule in these stories.

### Output

A `.sql` file shaped like this (illustrative, 3 rows):

```sql
 SELECT [RowId], [HTSNum], [Chapter99], [CountryofOrigin], [StartEffDate], [EndEffDate], [TariffType], [TariffGroup], [RequiredStatusCode], [ValidationLevel], [ExportDate] INTO [dbo].[tmdHTSAdditional]
	 FROM (VALUES
		(N'850132', N'99039466', N'TW', CAST(N'2026-05-01 00:00:00' AS DATETIME), CAST(N'9999-12-31 23:59:59' AS DATETIME), N'232', N'Exclusion', N'P', N'W', N'')
		,(N'850133', N'99039466', N'TW', CAST(N'2026-05-01 00:00:00' AS DATETIME), CAST(N'9999-12-31 23:59:59' AS DATETIME), N'232', N'Exclusion', N'P', N'W', N'')
		,(N'850134', N'99039466', N'TW', CAST(N'2026-05-01 00:00:00' AS DATETIME), CAST(N'9999-12-31 23:59:59' AS DATETIME), N'232', N'Exclusion', N'P', N'W', N'');
```

This is a **scaffold, not directly executable SQL** — see
[Caveats](#caveats--known-rough-edges). It is the raw material you embed into a
production deployment template.

---

## `tool.js` — post-process the generated rows

A dependency-free Node.js helper for working with a file of generated row tuples (the
"ROWS" file). Run `node tool.js --help` for built-in usage.

```
node tool.js <command> [options]
```

### `batch` — split rows into runnable `INSERT` batches

Splits the input into chunks of N rows and writes one `.sql` file per chunk. For each
chunk it rewrites the leading `,(N'` into a full `INSERT INTO <table> (<columns>) VALUES`
statement and terminates the chunk with `;`. Useful when a single statement would exceed
SQL Server's 1,000-row `VALUES` limit.

| Option | Default | Meaning |
|--------|---------|---------|
| `--input`, `-i` | `./ROWS` | Input rows file |
| `--amount`, `-a` | `712` | Rows per batch |
| `--output`, `-o` | `./batches` | Output directory |
| `--table`, `-t` | `[dbo].[tmdhtsAdditional]` | Table name for the `INSERT` |
| `--columns`, `-c` | the 10 contract columns | Comma-separated column list |

```bash
node tool.js batch -i ./ROWS -a 500 -o ./batches -t "[dbo].[tmdHTSAdditional]"
```

### `index` — add a `RowId` to each tuple

Inserts a sequential number as the first value of every tuple (`,(N'…'` → `,(1, N'…'`).
This lines each row up with the `[RowId]` column in `main.py`'s header line.

| Option | Default | Meaning |
|--------|---------|---------|
| `--input`, `-i` | `./ROWS` | Input rows file |
| `--output`, `-o` | `<input>_indexed` | Output file |

```bash
node tool.js index -i ./ROWS -o ./ROWS_indexed
```

### `merge` — concatenate files in a folder

Reads every file in a folder (e.g. the `batches/` directory) and concatenates them into
a single file, separated by newlines.

| Option | Default | Meaning |
|--------|---------|---------|
| `--folder`, `-f` | — (required) | Folder of files to merge |
| `--output`, `-o` | `<folder>_merged` | Output file |

```bash
node tool.js merge -f ./batches -o ./merged.sql
```

---

## End-to-end workflow

```
┌──────────────────────────────┐
│  Work-item Excel attachment   │   tmdhtsAdditional_*_Changes.xlsx
│  (InsertData tab, N rows)     │   — the approved record set
└───────────────┬──────────────┘
                │  main.py --excel … --sheet InsertData --out ROWS.sql
                ▼
┌──────────────────────────────┐
│  ROWS.sql                     │   header line + one ,(…) tuple per row
└───────────────┬──────────────┘
                │  (optional) tool.js index / batch / merge
                ▼
┌──────────────────────────────┐
│  Row tuples / INSERT batches  │
└───────────────┬──────────────┘
                │  hand-embed into the production template
                ▼
┌──────────────────────────────┐
│  V26.x.****__DATA_*.sql       │   backup + idempotent insert + NOLOCK +
│  (committed to gtm-legacy_    │   transaction + error handling
│   gtm-sql)                    │
└──────────────────────────────┘
```

A **verification script** (`VERIFY_*.sql`) is authored separately to check the work
item's acceptance criteria after deployment.

---

## What these tools do *not* do

Be aware of the boundaries — the committed deployment script for US 5441030 contains a
lot of logic that **none** of these tools generate:

- **No backup table** creation (story requirement R1).
- **No idempotency / existence check** — the production script inserts only
  `WHERE NOT EXISTS (…)` on `(HTSNum, Chapter99, TariffType, CountryofOrigin, …)`. The
  generated rows are a plain value list with no such guard.
- **No transaction, `TRY/CATCH`, or `NOLOCK`.**
- **No `NULL` handling for `ExportDate`** — the story says "ExportDate = NULL (leave
  blank)", but `main.py` emits `N''` (empty string) for blanks, and the committed script
  uses a sentinel `'1900-01-01'`. Reconcile this manually.
- **It does not generate the `VERIFY_*.sql` verification script.** That script is a
  read-only set of acceptance-criteria checks (record counts per Chapter 99 heading, the
  `9999-12-31 23:59:59` datetime check, duplicate check, dot-in-HTSNum check, and a
  PASS/FAIL roll-up). Nothing in this workspace produces it — it is derived from the work
  item's Acceptance Criteria (AC-1…AC-7) and authored by hand / with assistance per story.

In short: **these tools generate the data; a human still assembles and reviews the
production and verification SQL.**

---

## Caveats / known rough edges

The scripts are pragmatic utilities, not hardened software. Observed behaviors to be
aware of:

1. **`main.py` output is not directly executable.** The header is
   `SELECT … FROM (VALUES …` but the generated text never closes the derived table with
   the required `) AS T(<columns>)` alias, and the `SELECT` lists eleven columns
   (`[RowId]` + the ten) while each tuple has only ten values. You must either run
   `tool.js index` to add the `RowId` value **and** add the closing alias, or (more
   commonly) lift the tuples into a different template. Treat the output as a scaffold.
2. **First-comma / last-row logic depends on the DataFrame index.** `main.py` decides
   "first row" and "last row" by comparing the pandas index label to `0` and
   `len(df) - 1`. After `drop_duplicates()` (or with a non-default index) the surviving
   rows may not include those exact labels, so the leading-comma strip and the trailing
   `;` may not be applied as expected. Eyeball the first and last lines of the output.
3. **Blanks become `N''`, not `NULL`.** If a truly `NULL` value is required, edit it in.
4. **Date columns must carry their time component in Excel.** `EndEffDate` must be
   `9999-12-31 23:59:59`; if the spreadsheet stores it as a date-only value it becomes
   `00:00:00`, which fails the acceptance criteria. `main.py` preserves whatever time the
   cell has — it does not force `23:59:59`.
5. **`tool.js` joins lines without re-adding newlines** (`split("\n")` then `join("")`),
   so batch/index output is whitespace-compact rather than one-tuple-per-line. This is
   cosmetically different but still valid SQL.
6. The fallback date-detection branch in `sql_literal` uses `col.find("Date")`
   (truthy/falsy) rather than an explicit `!= -1` test; it only affects exotic
   non-string/non-numeric cell types and rarely fires in practice.

---

## Quick reference

```bash
# 1. Generate the row tuples from the approved spreadsheet
python main.py --excel Changes.xlsx --sheet InsertData --out ROWS.sql

# 2a. (optional) add RowId values
node tool.js index -i ./ROWS.sql -o ./ROWS_indexed.sql

# 2b. (optional) split into <=1000-row INSERT batches
node tool.js batch -i ./ROWS.sql -a 500 -o ./batches -t "[dbo].[tmdHTSAdditional]"

# 2c. (optional) merge the batches back into one file
node tool.js merge -f ./batches -o ./merged.sql

# 3. Hand-embed the rows into the production deployment template, then
#    author/run the VERIFY_*.sql verification script.
```
