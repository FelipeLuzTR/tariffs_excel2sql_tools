# Design: Generic Regulatory DBA Data-Script Generator

**Status:** Proposal / for review
**Author:** drafted with Claude (Opus 4.8), from analysis of US 5462916, US 5463147, and their merged PRs
**Scope:** Automate the generation of idempotent DBA data-maintenance SQL (deploy + verification + QA rollback dev-test) for regulatory "insert/update/delete reference data" stories, from a single standardized Excel workbook.

---

## 1. Problem & goal

GTM regularly ships **regulatory DBA data scripts**: a CBP message (e.g. CSMS 68855869 / Proclamation 11032) requires reference rows to be inserted/updated/deleted in a table such as `tmdHTSAdditional` or `tmgGlobalCodes`. Today each story is hand-built:

- An analyst attaches an **Excel workbook** of the required rows.
- A developer hand-writes an **idempotent SQL script** (backup → changes → idempotency guards) plus runs **manual AC verification**.

This is slow, repetitive, and error-prone. Two real defects in this very work illustrate the cost:
- **AC-6 verification bug** — counting *all* rows in a heading instead of only the payload (false FAIL on any DB with pre-existing data).
- **Defect 5463196** — the INSERT existence key omitted `StartEffDate`, so a new period record was skipped because an *expired* record matched the coarser key.

**Goal:** one **config-driven generator** + one **standardized input workbook** that produces the deploy script, the verification script, and a transaction+rollback QA dev-test — covering both the simple (`tmgGlobalCodes`) and complex (`tmdHTSAdditional`) shapes, and turning the lessons above into built-in safeguards.

---

## 2. Reference stories (the two anchors)

| | **US 5463147 — tmgGlobalCodes** (minimal) | **US 5462916 — tmdHTSAdditional** (full) |
|---|---|---|
| Target | `dbo.tmgGlobalCodes` (flat code/decode lookup) | `dbo.tmdHTSAdditional` (period-based, eff-date ranges) |
| Operations | **INSERT only** (6 rows) | **DELETE → UPDATE StartEff → UPDATE EndEff → INSERT** (ordered; ~51 / 17 / 179 / 1955 rows) |
| Columns | PartnerID, EffDate, FieldName, Code, Decode, StaticFlag, DeletedFlag, KeepDuringRollback | HTSNum, Chapter99, CountryofOrigin, StartEffDate, EndEffDate, TariffType, TariffGroup, RequiredStatusCode, ValidationLevel, ExportDate |
| Match / existence key | `(PartnerID, FieldName, Code)` | INSERT `(HTSNum, Chapter99, TariffType, CountryofOrigin, StartEffDate)`; UPDATEs and DELETE use narrower op-specific keys + value predicates |
| Idempotency | `MERGE … WHEN NOT MATCHED` | `INSERT … WHERE NOT EXISTS`; guarded `UPDATE … FROM … JOIN`; pattern `DELETE` |
| Special rules | PartnerID sourced from `tmfDefaults` (never hardcoded, R1); `Decode = Code`; constants `Y/N/N`; never delete (R5); dual-format codes (`99038223` **and** `9903.82.23`) | blank `HTSNum`/`CountryofOrigin` stored as `''` with `ISNULL(...,'')` normalization; targeted (pattern) delete of malformed rows |
| Deliverable PR | `tr/gtm-legacy_gtm-sql` #726 → `V26.2.0714__DATA_tmgGlobalCodes_232_Metals_CSMS68855869_5463147.sql` | `tr/gtm-legacy_gtm-sql` #727 → `V26.2.0713__DATA_tmdHTSAdditional_232_Metals_CSMS68855869.sql` |

**Both share ~80% identical skeleton** (next section). The differences are all *data/config*, not structure.

---

## 3. The unifying model

> **A run = a target table + an ordered list of operations.
> Each operation = (action data tab → op-type → match-key → idempotency-strategy → value-mapping → from-predicate).
> Everything is wrapped in a fixed safety skeleton, and the same metadata drives auto-generated verification.**

`tmgGlobalCodes` is this model with **one** INSERT operation; `tmdHTSAdditional` is the model with **four** ordered operations. The engine doesn't care how many.

---

## 4. Standardized input — one Excel workbook

The workbook carries both the **data** and a machine-readable **spec** (so the engine needs no per-story code). Control sheets are prefixed `_`.

### 4.1 `_Meta` (run identity) — key/value rows

| Key | Example (tmgGlobalCodes) | Example (tmdHTSAdditional) |
|---|---|---|
| `TargetTable` | `dbo.tmgGlobalCodes` | `dbo.tmdHTSAdditional` |
| `StoryId` | `5463147` | `5462916` |
| `Feature` | `232_Metals_CSMS68855869` | `232_Metals_CSMS68855869` |
| `Release` | `26.2` | `26.2` |
| `EffectiveDate` | `2026-06-08 00:00:00` | `2026-06-08 00:00:00` |
| `RetireDate` *(opt)* | — | `2026-06-07 23:59:59` |
| `BackupSchema` | `bck` | `bck` |
| `PartnerScoped` | `Y` | `N` |
| `PartnerSource` | `SELECT TOP 1 PartnerID FROM dbo.tmfDefaults WITH (NOLOCK)` | — |
| `NeverDelete` *(business rule guard)* | `Y` | `N` |

Backup name is derived from a fixed convention (§7): `[bck].[bck_<Table>_<Feature>_<StoryId>]` (deterministic — no generation-date stamp, so the same workbook always yields byte-identical SQL).

### 4.2 `_Columns` (one row per target column)

| Field | Meaning |
|---|---|
| `ColumnName` | target column |
| `SqlType` | type for the staging table variable (`varchar(30)`, `datetime`, …) |
| `Source` | how the value is produced (grammar below) |
| `NullNormalize` | `Y` → wrap with `ISNULL(col,'')` in match predicates |

**`Source` grammar:**
- `CELL` — value from the action-tab cell of the same column
- `CONST:<v>` — literal constant (e.g. `StaticFlag = CONST:Y`)
- `ECHO:<OtherColumn>` — copy another column (e.g. `Decode = ECHO:Code`)
- `PARAM:<name>` — a script-level parameter declared in `_Meta` (`EffDate = PARAM:EffectiveDate`, `PartnerID = PARAM:PartnerID`)
- `NULL` — literal `NULL`

### 4.3 `_Operations` (the heart — one row per operation)

| Field | Meaning |
|---|---|
| `Order` | execution order (backup is always first, implicit) |
| `ActionTab` | sheet holding this op's rows (blank if pure predicate) |
| `ActionFilter` | value in the tab's `Action` column that selects rows for this op (e.g. `Insert`); other values (`Already in prod`, `No Action`) are skipped |
| `OpType` | `INSERT` / `UPDATE` / `DELETE` |
| `MatchKey` | comma-separated columns = existence/join key |
| `FromPredicate` | SQL predicate selecting *existing* rows to act on (UPDATE/DELETE), e.g. `StartEffDate = '2026-04-06 00:00:00'`, or a malformed-row pattern |
| `SetMap` | for UPDATE: `Col <- Source` (e.g. `EndEffDate <- CELL(New_EndEffDate)` or `StartEffDate <- PARAM:EffectiveDate`) |
| `Idempotency` | `NOT_EXISTS` (insert) / `GUARDED` (update; from-predicate makes re-run a no-op) / `PATTERN` (delete by predicate; re-run deletes 0) |
| `VerifyGroupBy` | column(s) to group expected counts by for AC verification (e.g. `Chapter99`, `FieldName`) |

### 4.4 Action tabs & `Notes`
- **Action tabs** hold the data: **one row per record, one column per `CELL` field only** (`PARAM`/`CONST`/`ECHO`/`NULL` columns are supplied by the engine — not in the tab; extra columns are ignored). UPDATE tabs add the `New_<Col>` "to" value. Multi-format rows (e.g. `99038223` and `9903.82.23`) are simply **two data rows** — no engine logic. The `Action` column is **optional** and read **only** when `ActionFilter` is set (to skip non-actionable rows like `Already in prod`); the engine **warns** if an `Action` column is present but unused, so it can't quietly mislead. A pattern `DELETE` reads no tab (`ActionTab` blank).
- **`Notes`** (and any `notes_*` sheet) is free text, ignored by the engine.

### 4.5 Worked `_Operations` for the two stories

**tmgGlobalCodes:**
| Order | ActionTab | ActionFilter | OpType | MatchKey | Idempotency | VerifyGroupBy |
|---|---|---|---|---|---|---|
| 1 | INSERTS | Insert | INSERT | PartnerID,FieldName,Code | NOT_EXISTS | FieldName |

**tmdHTSAdditional:**
| Order | ActionTab | ActionFilter | OpType | MatchKey | FromPredicate | SetMap | Idempotency | VerifyGroupBy |
|---|---|---|---|---|---|---|---|---|
| 1 | Deletes | — | DELETE | — | `Chapter99='99038212' AND TariffType='232' AND ((ISNULL(HTSNum,'')='' AND ISNULL(CountryofOrigin,'') IN ('BY','CU','KP','RU')) OR (ISNULL(HTSNum,'')<>'' AND ISNULL(CountryofOrigin,'')=''))` | — | PATTERN | — |
| 2 | Update_StartEffDate | — | UPDATE | HTSNum,Chapter99,TariffType | `StartEffDate='2026-04-06 00:00:00'` | `StartEffDate <- CELL(New_StartEffDate)` | GUARDED | — |
| 3 | Update_EndEffDate | — | UPDATE | HTSNum,Chapter99,TariffType,CountryofOrigin | `EndEffDate='9999-12-31 23:59:59'` | `EndEffDate <- CELL(New_EndEffDate)` | GUARDED | — |
| 4 | Inserts_tmdhtsadditional | Insert | INSERT | HTSNum,Chapter99,TariffType,CountryofOrigin,StartEffDate | — | — | NOT_EXISTS | Chapter99 |

---

## 5. The generic engine — one tool, three outputs

`gen_dba_script.py <workbook.xlsx>` reads `_Meta` / `_Columns` / `_Operations`, then emits:

1. **Deploy script** — `V<release>.XXXX__DATA_<Table>_<Feature>_<StoryId>.sql`
2. **Verification script** — `VERIFY_<Table>_<Feature>_<StoryId>.sql`
3. **QA dev-test (rolls back)** — `DEVTEST_<Table>_<StoryId>_DONOTCOMMIT.sql`

It is the architecture already built in `generate_sql.py`, generalized: two shared builders compose all three outputs, so a fix lives in one place.

```
load_workbook(xlsx) -> spec(meta, columns, operations) + per-op data tuples
        │
        ├── data_ops_sql(spec)        # backup + each op block (ordered), idempotent
        ├── verification_sql(spec)    # AC checks derived from operations/data
        │
deploy  = preamble + BEGIN TRAN + data_ops_sql(spec) + COMMIT + diagnostics + R8 restore block
verify  = preamble + verification_sql(spec)
dev-test = preamble + BEGIN TRAN + data_ops_sql(spec) [+ re-run for idempotency] + verification_sql(spec) + ROLLBACK
```

---

## 6. The fixed SQL skeleton (always emitted)

Shared by both reference scripts; the engine emits it verbatim, parameterized only by table/columns:

- **Pre-flight guards** — `OBJECT_ID(<table>)`, `SCHEMA_ID('bck')`, and `OBJECT_ID('dbo.tmfDefaults')` when partner-scoped; each returns a diagnostic result set instead of throwing.
- **PartnerID resolution** (when `PartnerScoped=Y`) — `DECLARE @PartnerID INT = (<PartnerSource>);` + null guard. Never hardcoded (R1).
- **Idempotent backup** — `IF OBJECT_ID(@BackupName) IS NULL → SELECT <cols> INTO <backup> FROM <table> WITH (NOLOCK)`; never overwritten (re-run safe).
- **Single `BEGIN TRANSACTION` … `COMMIT`**, inside `BEGIN TRY/CATCH`; `ROLLBACK` when `@@TRANCOUNT>0`; captures `ERROR_MESSAGE/SEVERITY/STATE/LINE`.
- **`WITH (NOLOCK)`** on every SELECT.
- **Set-based staging** — each op loads a `@`table-variable of literal rows (chunked ≤ 1000/`INSERT…VALUES`) then runs one set operation.
- **Diagnostic output contract** — final `SELECT DatabaseName=DB_NAME(), <per-op counters>, BackupTable=@BackupName, Msgs=@Msg` where `@Msg` is a CRLF-built log.
- **R8 QA-only restore block** (commented) — `TRUNCATE` + `INSERT … SELECT * FROM <backup>` + `DROP <backup>`, with the "never in production" warning.

### Idempotency strategy per op
| OpType | Emitted form | Re-run behavior |
|---|---|---|
| INSERT | `INSERT … SELECT … FROM @stg s WHERE NOT EXISTS (… match on MatchKey, NullNormalize where flagged)` | inserts 0 |
| UPDATE | `UPDATE t SET <SetMap> FROM <table> t JOIN @stg k ON <MatchKey> WHERE <FromPredicate>` | updates 0 (rows no longer match the "from" value) |
| DELETE (pattern) | `DELETE FROM <table> WHERE <DeletePredicate>` | deletes 0 (malformed rows already gone) |
| DELETE (by-key) | `DELETE t FROM <table> t JOIN @stg k ON <MatchKey>` | deletes 0 |

---

## 7. Verification generation (AC, done right)

From `_Operations.VerifyGroupBy` + the action data, the engine emits:
- **Per-group counts scoped to the payload** — counts only rows that match the inserted keys (`Present_from_payload`), so unrelated pre-existing data never inflates the result. *(This is the AC-6 lesson — a blanket `COUNT(*) WHERE heading=X` is wrong on any non-empty DB.)*
- **"Exactly N" union check** across all inserted keys.
- **No-duplicate check** on the **full** existence key (including date columns where in the key). *(The 5463196 lesson — the key must distinguish period records.)*
- **UPDATE checks** — count of target rows now carrying the new value, scoped to the update keys.
- **DELETE checks** — malformed/target rows remaining = 0.
- **Re-run idempotency** — second application changes 0 rows.
- **PASS/FAIL roll-up** + notes where literal AC prose deviates from data reality (we hit this with AC-3 "0 before reinsert" and AC-4 blanket counts).

---

## 8. Built-in safeguards (lessons → features)

| Lesson from this work | Built-in safeguard |
|---|---|
| Wrong existence key skipped a new period record (defect 5463196) | **Key-duplicate lint:** fail generation if the chosen `MatchKey` yields duplicate rows within the action tab (would have surfaced the missing `StartEffDate`). |
| Verification counted pre-existing rows (AC-6) | Verification is **payload-scoped** by construction. |
| `@BrokenCount` declared twice (parse error only SQL Server caught) | **Duplicate-DECLARE lint** + balanced-parens check on every generated script. |
| Dev-test reimplemented logic → drift risk | **Single source of truth:** dev-test composes the *same* builders as deploy+verify. |
| AC prose sometimes conflicts with data | Engine emits the **corrected** scoped checks and **flags** the deviation, rather than matching prose blindly. |
| Deploy must not silently change | Generator supports a **byte-diff check** against a prior committed script. |

---

## 9. How we prove the engine is correct (acceptance test)

Author the standardized workbooks for the two anchor stories and **regenerate their deliverables**, then **diff against the merged, human-reviewed PRs**:
- tmgGlobalCodes → reproduce the substance of `V26.2.0714` (PR #726)
- tmdHTSAdditional → reproduce `V26.2.0713` (PR #727)

If the engine reproduces two independently-authored, reviewed scripts (modulo cosmetic formatting), it is trustworthy. This is the generic tool's own acceptance criterion.

---

## 10. Limits & non-goals (where human judgment stays)

- **Match-key selection is a human decision.** The tool lints for duplicates but cannot *know* that period records need `StartEffDate` in the key. The format makes the key **explicit and reviewable**; it does not remove the judgment.
- **Pattern deletes** (e.g. the malformed-`9903.82.12` cleanup) cannot be inferred — they require an explicit `DeletePredicate`, supplied by someone who understands the data defect.
- **Business rules** (`NeverDelete`, "PartnerID never hardcoded") are config flags/guards — someone must set them correctly.
- **The Excel still requires a competent analyst.** Garbage rows in → garbage SQL out. The standard reduces structural ambiguity, not domain correctness.
- **Not in scope:** schema/DDL changes, application/config/code changes, cross-table transactional logic beyond a single target table per workbook.

---

## 11. File & PR conventions (carried from the merged PRs)

- **Path:** `Database/Application/<release>/Hotfix/V<release>.<seq>__DATA_<Table>_<Feature>_<CSMS>_<StoryId>.sql` (the `<seq>` is assigned at integration — PR #726's `0712` was renumbered to `0714`; the generator emits an `XXXX` placeholder).
- **Backup name:** `[bck].[bck_<Table>_<Feature>_<StoryId>]` (deterministic; `StoryId` makes it unique without a date stamp).
- **PR flow:** feature branch → `develop`, then a cherry-pick PR into `release/<n>` (`AB#<StoryId>` in title). Repo default branch is **`develop`**, not `master`.

---

## 12. Phased plan

1. **Lock the contract** — finalize the `_Meta` / `_Columns` / `_Operations` schema (this doc) and the SQL skeleton (already battle-tested in `generate_sql.py`).
2. **Build the engine** — generalize `generate_sql.py` into `gen_dba_script.py` (workbook → 3 outputs) with the value-source grammar and the per-op idempotency strategies.
3. **Prove it** — regenerate both anchor deliverables; diff vs PRs #726/#727.
4. **Harden** — add the lints (key-duplicate, duplicate-DECLARE, parens), the verification generator, and the QA dev-test composition.
5. **Document & template** — ship a blank standardized workbook + the two anchors as worked examples + a README and the conventions above.

---

## 13. Open questions for PO / team

1. **Existence-key as a requirement field.** Should the work-item template explicitly require the match key per operation (incl. date columns)? Defect 5463196 shows leaving it implicit is dangerous.
2. **Idempotency mechanism standard.** `MERGE` (tmgGlobalCodes) vs `INSERT…WHERE NOT EXISTS` (tmdHTSAdditional) are equivalent — standardize on one for all generated scripts? (Recommend `INSERT…WHERE NOT EXISTS` + guarded UPDATE/DELETE — simpler to read and review.)
3. **Verification artifact.** tmgGlobalCodes shipped with no separate verify file; tmdHTSAdditional did. Standardize on always emitting the verify script + QA dev-test?
4. **Where the tool lives.** This repo (`tariffs_excel2sql_tools`) generalizing `main.py`/`tool.js`/`generate_sql.py`, or a home in the GTM SQL tooling?
5. **AC prose vs. generated checks.** When literal AC wording conflicts with data reality, is it acceptable for the generator to emit the corrected check + a flagged note (current approach), or must AC wording be amended first?
