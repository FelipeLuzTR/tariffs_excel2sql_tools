# Standardized workbook contract (reference)

A workbook = three control sheets (prefixed `_`) describing the change, plus action data tabs.
This is the input contract for the `gen-dba-script` generator. Authoritative copy:
`dba-script-generator/DESIGN.md`.

## `_Meta` (key/value rows)

| Key | Notes |
|---|---|
| `TargetTable` | e.g. `dbo.tmgGlobalCodes` |
| `StoryId` | ADO work-item id |
| `Feature` | slug, e.g. `232_Metals_CSMS68855869` |
| `Release` | e.g. `26.2` (drives the deploy filename) |
| `EffectiveDate` | new records' start date |
| `RetireDate` | (optional) end-date used by expiry UPDATEs |
| `BackupSchema` | usually `bck` |
| `PartnerScoped` | `Y` if the table has a `PartnerID` |
| `PartnerSource` | SQL to resolve `@PartnerID` (e.g. `SELECT TOP 1 PartnerID FROM dbo.tmfDefaults WITH (NOLOCK)`) |
| `NeverDelete` | business-rule flag (e.g. `Y` for `tmgGlobalCodes`) |

## `_Columns` (one row per target column)

`ColumnName | SqlType | Source | NullNormalize`

Value-source grammar:

| `Source` | Meaning |
|---|---|
| `CELL` | value comes from the action-tab cell of the same name |
| `CONST:<v>` | literal constant (e.g. `CONST:Y`) |
| `ECHO:<Col>` | copy another column (e.g. `Decode = ECHO:Code`) |
| `PARAM:<name>` | a script parameter (`PARAM:EffectiveDate`, `PARAM:PartnerID`) |
| `NULL` | literal `NULL` |

`NullNormalize = Y` wraps the column in `ISNULL(col,'')` for match predicates (blank-safe).

## `_Operations` (one row per operation, executed in `Order`)

| Field | Meaning |
|---|---|
| `Order` | execution order (backup is always first, implicit) |
| `ActionTab` | data sheet for this op; **blank for a pattern DELETE** (it needs no data) |
| `ActionFilter` | *(optional)* value in an `Action` column to select rows (e.g. `Insert`); used only to skip non-actionable reference rows |
| `OpType` | `INSERT` / `UPDATE` / `DELETE` |
| `MatchKey` | comma-separated columns = existence / join key |
| `FromPredicate` | which existing rows to act on (UPDATE/DELETE) |
| `SetMap` | for UPDATE: `Col <- CELL(New_Col)` or `Col <- PARAM:x` |
| `Idempotency` | `NOT_EXISTS` (insert) / `GUARDED` (update; from-predicate makes re-run a no-op) / `PATTERN` (predicate delete) |
| `VerifyGroupBy` | column to group expected counts by (e.g. `Chapter99`, `FieldName`) |

## Data tabs — keep them lean

- One row per record, **one column per `CELL` field only**. `PARAM`/`CONST`/`ECHO`/`NULL` columns are supplied by the engine — do not put them in the tab. Extra columns are ignored.
- The `Action` column is optional and read **only** when `ActionFilter` is set. On a single-purpose tab, omit it. (The generator warns if an `Action` column is present but unused.)

## Match-key guidance (the highest-risk decision)

The match key must **uniquely identify a record**. For **period-based** tables (effective-date
ranges), include the date column (e.g. `StartEffDate`) — otherwise a new period record collides
with an expired one and the insert is silently skipped (a real, shipped defect, US 5463196).
The generator lints for duplicate keys within the payload, but choosing the key is human judgment.
