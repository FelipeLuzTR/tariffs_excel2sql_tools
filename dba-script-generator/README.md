# DBA Script Generator

Generate the three artifacts for a regulatory **DBA data-maintenance** story —
an idempotent **deploy** script, a read-only **verification** script, and a
transaction+rollback **QA dev-test** — from **one standardized Excel workbook**.

It turns the repetitive, error-prone hand-writing of `tmdHTSAdditional` /
`tmgGlobalCodes` insert/update/delete scripts into a config-driven, repeatable step.

---

## TL;DR

**The concept:** describe a regulatory data change **once** in a single Excel workbook —
*which table, which columns, which operations* — and the tool writes the production SQL
for you. No more hand-writing backup / insert / update / delete scripts (and no more bugs
from getting the existence key or the verification wrong).

**Where the workbook comes from — you author it.** A regulatory story arrives with the
analyst's spreadsheet of the rows to change; you express it in this tool's standard format.
The fastest way: **copy a starting point** and fill it in —
- `samples/TEMPLATE.xlsx` — a blank, structured starter, **or**
- `samples/STD_tmgGlobalCodes_5463147.xlsx` — the simplest worked example to clone.

A workbook has three control sheets (`_Meta`, `_Columns`, `_Operations`) that describe the
change, plus your data tabs. **The workbook *is* the requirements, in machine-readable form.**
(Full contract in §5; if you already have the analyst's work-item spreadsheet, you wrap it —
see §2.)

**Then generate + run, in 3 steps:**

```bash
# 1. install
pip install pandas openpyxl

# 2. generate all three artifacts (production-ready, release-aware names) into a folder
python gen_dba_script.py --workbook samples/STD_tmgGlobalCodes_5463147.xlsx --out-dir out/

# 3. dry-run on QA, then deploy for real
#    • run the DEVTEST_*.sql on QA   → applies + verifies + ROLLS BACK; expect every roll-up column = PASS
#    • run the V<rel>.XXXX__DATA_*.sql → the actual deployment (idempotent, backup-first)
#    • run the VERIFY_*.sql afterward  → post-deploy sign-off
```

Step 2 writes, into `out/`:

| File (derived from `_Meta`) | What it is |
|------|-----------|
| `V<Release>.XXXX__DATA_<Table>_<Feature>_<StoryId>.sql` | the real change — **idempotent**, backup-first, single transaction, safe to re-run. `XXXX` is the Flyway sequence you assign at integration. |
| `VERIFY_<Table>_<Feature>_<StoryId>.sql` | read-only **acceptance checks** (payload-scoped) for post-deploy sign-off |
| `DEVTEST_<Table>_<StoryId>_DONOTCOMMIT.sql` | a **QA dry-run**: applies everything, verifies it, then `ROLLBACK`s — nothing is saved |

Every generated SELECT uses `WITH (NOLOCK)`; counts are **payload-scoped** so pre-existing rows never inflate them.

---

## 1. Prerequisites

```bash
pip install pandas openpyxl       # Python 3.9+
```

## 2. Getting a workbook

You produce the standardized workbook one of two ways:

1. **From a starter (recommended).** Copy `samples/TEMPLATE.xlsx` (blank) or a sample
   (`samples/STD_*.xlsx`), then fill in the three control sheets + your data tabs per §5.
2. **Wrap the work-item spreadsheet.** If the analyst already provided the rows in their own
   Excel, keep those as the data tabs and add the `_Meta` / `_Columns` / `_Operations` control
   sheets that describe them. `make_samples.py` shows exactly this — it builds the
   `tmdHTSAdditional` sample by wrapping the work-item spreadsheet:
   ```bash
   python make_samples.py --hts-source <CSMS..._FINAL.xlsx>
   ```

There is no fully-automatic step here: the workbook captures human decisions (which rows,
which match key, which operations). The tool removes the *SQL-writing* drudgery, not the
domain judgment.

## 3. Run it

```bash
python gen_dba_script.py --workbook <your-workbook.xlsx> --out-dir out/
```

`--out-dir` writes all three artifacts with the production-ready names above.
For one-off control you can instead pass explicit paths: `--out`, `--out-verify`, `--out-test`.

## 4. Use the outputs

1. **`DEVTEST_*.sql` — dry-run on QA first.** Run it in SSMS against the target QA DB. It applies the real operations, re-runs them to prove idempotency (PASS 2 = all 0), runs the AC verification against the *uncommitted* state, then **`ROLLBACK`s — nothing is persisted.** Every roll-up column should read `PASS`.
2. **`V<Release>.XXXX__DATA_*.sql` — the real deployment.** Backup-first, idempotent, single transaction, safe to re-run. Set the `XXXX` Flyway sequence and drop it in the SQL repo (`Database/Application/<rel>/Hotfix/`).
3. **`VERIFY_*.sql` — post-deploy sign-off.** Read-only AC checks; run after deploy and confirm every roll-up column reads `PASS`.

---

## 5. The workbook contract (quick reference)

Control sheets (prefixed `_`) carry the spec; the rest are data.

**`_Meta`** — key/value rows: `TargetTable`, `StoryId`, `Feature`, `Release`,
`EffectiveDate`, `RetireDate` (opt), `BackupSchema`, `PartnerScoped` (Y/N),
`PartnerSource` (SQL to resolve `@PartnerID`), `NeverDelete`.

**`_Columns`** — one row per target column: `ColumnName | SqlType | Source | NullNormalize`.

| `Source` | Meaning |
|---|---|
| `CELL` | value from the action-tab cell of the same name |
| `CONST:<v>` | literal constant (e.g. `CONST:Y`) |
| `ECHO:<Col>` | copy another column (e.g. `Decode = ECHO:Code`) |
| `PARAM:<name>` | script parameter (`PARAM:EffectiveDate`, `PARAM:PartnerID`) |
| `NULL` | literal `NULL` |

`NullNormalize = Y` wraps the column in `ISNULL(col,'')` in match predicates.

**`_Operations`** — one row per operation (executed in `Order`):

| Field | Meaning |
|---|---|
| `Order` | execution order (backup is always first) |
| `ActionTab` | data sheet for this op |
| `ActionFilter` | value in the tab's `Action` column selecting rows (e.g. `Insert`) |
| `OpType` | `INSERT` / `UPDATE` / `DELETE` |
| `MatchKey` | comma-separated columns = existence / join key |
| `FromPredicate` | which existing rows to act on (UPDATE/DELETE) |
| `SetMap` | for UPDATE: `Col <- CELL(New_Col)` or `Col <- PARAM:x` |
| `Idempotency` | `NOT_EXISTS` / `GUARDED` / `PATTERN` |
| `VerifyGroupBy` | column to group expected counts by (e.g. `Chapter99`) |

> **The match key is the critical field.** It must uniquely identify a record —
> for period-based tables include the date column (e.g. `StartEffDate`).
> Getting it wrong is how a real defect (US 5463196) slipped a missing insert
> past review. The engine lints for duplicate keys within the data, but the
> *choice* is human judgment.

---

## 6. Worked examples (in `samples/`)

| Workbook | Story | Shape |
|---|---|---|
| `samples/TEMPLATE.xlsx` | — | blank, structured starter to copy |
| `samples/STD_tmgGlobalCodes_5463147.xlsx` | US 5463147 | 1 INSERT, partner-scoped, `Decode=Code`, dual-format codes |
| `samples/STD_tmdHTSAdditional_5462916.xlsx` | US 5462916 | DELETE + 2 UPDATEs + INSERT, period-based, blank-COO normalization |

Pre-generated outputs are in [`samples/out/`](samples/out/). Regenerate them with:

```bash
python make_samples.py                                   # rebuilds TEMPLATE + the tmgGlobalCodes sample
python gen_dba_script.py --workbook samples/STD_tmgGlobalCodes_5463147.xlsx --out-dir samples/out
python gen_dba_script.py --workbook samples/STD_tmdHTSAdditional_5462916.xlsx --out-dir samples/out
```

---

## 7. What it does *not* decide for you

- **The match key** (see warning above) — explicit and reviewable, not inferred.
- **Pattern deletes** (cleaning up malformed rows) need an explicit `FromPredicate`.
- **Business rules** (`NeverDelete`, "PartnerID never hardcoded") are config flags you set.
- Garbage data in → garbage SQL out. The standard removes structural ambiguity, not domain correctness.

---

## 8. Provenance

Validated against two already-shipped, human-reviewed deliverables
(`V26.2.0713` / `V26.2.0714` in `gtm-legacy_gtm-sql`): the generated
`tmdHTSAdditional` deploy is **data-identical** to the QA-validated script, and
the `tmgGlobalCodes` deploy is **semantically equivalent** to the merged PR. See
[DESIGN.md](DESIGN.md) for the full background, the unifying model, and limits.
