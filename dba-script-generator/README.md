# DBA Script Generator

Generate the three artifacts for a regulatory **DBA data-maintenance** story —
an idempotent **deploy** script, a read-only **verification** script, and a
transaction+rollback **QA harness** — from **one standardized Excel workbook**.

It turns the repetitive, error-prone hand-writing of `tmdHTSAdditional` /
`tmgGlobalCodes` insert/update/delete scripts into a config-driven, repeatable step.

- **Input:** one `.xlsx` with control sheets (`_Meta`, `_Columns`, `_Operations`) + action data tabs.
- **Output:** `*.deploy.sql`, `*.verify.sql`, `*.harness.sql`.
- One engine covers both the simplest (single INSERT) and the most complex (ordered DELETE + 2 UPDATEs + INSERT) shapes. Full design rationale in [DESIGN.md](DESIGN.md).

---

## 1. Prerequisites

```bash
pip install pandas openpyxl       # Python 3.9+
```

## 2. Run it

```bash
python gen_dba_script.py \
    --workbook   samples/STD_tmgGlobalCodes_5463147.xlsx \
    --out        out.deploy.sql \
    --out-verify out.verify.sql \
    --out-test   out.harness.sql
```

`--out` is required; `--out-verify` and `--out-test` are optional.

## 3. Use the outputs

1. **`*.harness.sql` — dry-run on QA first.** Run it in SSMS against the target QA DB. It applies the real operations, re-runs them to prove idempotency (PASS 2 = all 0), runs the AC verification against the *uncommitted* state, then **`ROLLBACK`s — nothing is persisted.** Every roll-up column should read `PASS`.
2. **`*.deploy.sql` — the real deployment.** Backup-first, idempotent, single transaction. Safe to re-run. Assign the Flyway version number and drop it in the SQL repo (`Database/Application/<rel>/Hotfix/`).
3. **`*.verify.sql` — post-deploy sign-off.** Read-only AC checks; run after deploy and confirm every roll-up column reads `PASS`.

Every generated SELECT uses `WITH (NOLOCK)`; counts are **payload-scoped** so pre-existing rows in a heading never inflate them.

---

## 4. The workbook contract (quick reference)

Control sheets (prefixed `_`) carry the spec; the rest are data.

**`_Meta`** — key/value rows: `TargetTable`, `StoryId`, `Feature`, `Release`,
`EffectiveDate`, `RetireDate` (opt), `BackupSchema`, `PartnerScoped` (Y/N),
`PartnerSource` (SQL to resolve `@PartnerID`), `NeverDelete`.

**`_Columns`** — one row per target column: `ColumnName | SqlType | Source | NullNormalize`.
`Source` grammar:

| Source | Meaning |
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

## 5. Worked examples (in `samples/`)

| Workbook | Story | Shape |
|---|---|---|
| `samples/STD_tmgGlobalCodes_5463147.xlsx` | US 5463147 | 1 INSERT, partner-scoped, `Decode=Code`, dual-format codes |
| `samples/STD_tmdHTSAdditional_5462916.xlsx` | US 5462916 | DELETE + 2 UPDATEs + INSERT, period-based, blank-COO normalization |

Pre-generated outputs are in [`samples/out/`](samples/out/).

Rebuild the simple (self-contained) sample and regenerate its outputs:

```bash
python make_samples.py
python gen_dba_script.py --workbook samples/STD_tmgGlobalCodes_5463147.xlsx \
    --out samples/out/tmgGlobalCodes.deploy.sql \
    --out-verify samples/out/tmgGlobalCodes.verify.sql \
    --out-test samples/out/tmgGlobalCodes.harness.sql
```

The complex sample was built from the work-item spreadsheet:
`python make_samples.py --hts-source <CSMS..._FINAL.xlsx>`.

---

## 6. What it does *not* decide for you

- **The match key** (see warning above) — explicit and reviewable, not inferred.
- **Pattern deletes** (cleaning up malformed rows) need an explicit `FromPredicate`.
- **Business rules** (`NeverDelete`, "PartnerID never hardcoded") are config flags you set.
- Garbage data in → garbage SQL out. The standard removes structural ambiguity, not domain correctness.

---

## 7. Provenance

Validated against two already-shipped, human-reviewed deliverables
(`V26.2.0713` / `V26.2.0714` in `gtm-legacy_gtm-sql`): the generated
`tmdHTSAdditional` deploy is **data-identical** to the QA-validated script, and
the `tmgGlobalCodes` deploy is **semantically equivalent** to the merged PR. See
[DESIGN.md](DESIGN.md) for the full background, the unifying model, and limits.
