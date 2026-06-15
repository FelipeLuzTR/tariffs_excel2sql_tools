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

**Where the workbook comes from.** The BA/SME provides the regulatory **story + a data
spreadsheet** (the rows, counts, effective dates) — that is their job, and it's complete as-is.
An **AI skill drafts the standardized workbook** from it (`csms-to-dba-script` from a CBP
bulletin, or `build_workbook.py` from the BA's attachment — §2), and the **developer reviews**
the result at the approval gate (§3) before any SQL is written. You can also author/clone a
workbook by hand (`samples/TEMPLATE.xlsx`).

Either way, a workbook is three control sheets (`_Meta`, `_Columns`, `_Operations`) + data
tabs — **the requirements in machine-readable form** (full contract in §5).

**Then generate + run, in 3 steps:**

```bash
# 1. install
pip install pandas openpyxl

# 2. review the operations (no files written), then generate after a human approves
python gen_dba_script.py --workbook samples/STD_tmgGlobalCodes_5463147.xlsx --out-dir out/                       # review only
python gen_dba_script.py --workbook samples/STD_tmgGlobalCodes_5463147.xlsx --out-dir out/ --confirm-operations  # writes

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

## Who does what (and where the AI skills fit)

This replaces "the Dev hand-writes the SQL." The Dev now **reviews** the AI's translation instead of authoring it — nobody hand-writes the script, and nobody rubber-stamps it either.

| Role | Owns | Does |
|---|---|---|
| **BA / SME** | regulatory **intent + data** | writes the story (operations, R-requirements, effective dates) and attaches the data spreadsheet. Does **not** write SQL or pick match keys — that's technical, and they shouldn't be asked to. |
| **AI skills** | the **translation** | draft the standardized workbook and generate the SQL — the drudgery the Dev used to do by hand. Never deploy; always stop at the gate. |
| **Developer** | the **engineering review** | reviews the AI's operation interpretation (match keys + guards) at the approval gate, ratifies the judgment calls, approves. Owns the dev-test + deploy. |

The pieces:

- **`csms-to-dba-script`** (skill) — interprets a CBP bulletin → proposes a workbook (bulletin-driven changes).
- **`build_workbook.py`** + a per-table profile + a per-story spec — turns a BA's raw delta attachment → a workbook (§2).
- **`gen-dba-script`** (skill / `gen_dba_script.py`) — turns a workbook → deploy/verify/dev-test SQL, **behind the operations-approval gate** (§3).

**End-to-end:** SME story + data → AI drafts workbook → AI generates SQL *(gate stops it)* → **Dev reviews + approves** → dev-test on QA → deploy → verify.

---

## 1. Prerequisites

```bash
pip install pandas openpyxl       # Python 3.9+
```

## 2. Getting a workbook

You produce the standardized workbook one of three ways:

1. **From a starter.** Copy `samples/TEMPLATE.xlsx` (blank) or a sample (`samples/STD_*.xlsx`),
   then fill in the three control sheets + your data tabs per §5.
2. **From a BA attachment, via the adapter (recommended — repeatable + reviewable).** If the
   analyst/BA attached a raw delta spreadsheet, don't hand-write a one-off build script — use
   the committed, config-driven adapter `build_workbook.py`. It pairs a per-table **profile**
   (`profiles/<table>.json` — the `_Columns`, column aliases, and default match keys) with a
   small per-story **spec** (`.json` — the operations: which source tab → which op, match key,
   guard, set):
   ```bash
   python build_workbook.py --attachment <BA_delta.xlsx> --spec <story.json> --out <workbook.xlsx>
   ```
   The spec is the engineering interpretation in diffable, reviewable form. Worked example:
   `specs/EXAMPLE_5475122_tmdHTSAdditional.json` — a 4-op remediation (keyed DELETE + two
   guarded UPDATEs + INSERT). Same attachment + spec → same workbook, every time.
3. **Wrap it by hand.** Keep the analyst's rows as data tabs and add the control sheets
   yourself; `make_samples.py --hts-source <FINAL.xlsx>` shows this pattern.

There is no fully-automatic step here: the **spec** (or the control sheets) captures human
decisions — which rows, which match key, which operations. The tools remove the *SQL-writing*
and *assembly* drudgery, not the domain judgment. The match keys + guards in the spec are the
engineering layer, and `gen_dba_script.py` gates on them (§3) before any SQL is written.

## 3. Run it — review, THEN confirm

The generator **writes nothing without `--confirm-operations`** (an approval gate). Run it twice:

```bash
# 1. Review: prints the _Operations interpretation (match keys, guards, idempotency). No files.
python gen_dba_script.py --workbook <your-workbook.xlsx> --out-dir out/

# 2. Generate: only after a human has approved that interpretation.
python gen_dba_script.py --workbook <your-workbook.xlsx> --out-dir out/ --confirm-operations
```

Why the gate: the action-tab **data** is BA/SME-sourced, but `_Meta`/`_Columns`/`_Operations`
(the **match keys and guards**) are the engineering-judgment layer. When those were AI-authored,
they must be human-reviewed first — the dev-test validates the *data*, not whether the *key* is
right in principle. `--out-dir` writes all three artifacts with the production-ready names above;
for one-off control pass explicit `--out`, `--out-verify`, `--out-test`.

### What the developer checks at the gate

The gate prints each operation (match key, guard, idempotency). You confirm it against the
**story — which is your answer key** — *not* by trusting the AI's "looks faithful" summary:

1. **Each op's match key + guard matches the story's Script Requirements** (R3 existence keys, R8 scope/order, R9 delete guard). A mechanical cross-check, op by op.
2. **The counts match** the spreadsheet / Summary tab.
3. **Ratify the judgment calls** — the few places the AI went *beyond* the literal spec and you must actively decide:
   - a **period-safe match key** (e.g. adding `StartEffDate` per defect 5463196 even when the AC text lists fewer columns);
   - translating prose like *"rows currently open"* into a concrete **guard predicate** (e.g. `EndEffDate > '<date>'`).

The **SME is not asked to confirm this** — they already provided the story + data; the SQL mechanics are the Dev's call. And the dev-test that follows validates counts + idempotency on real data but **not** whether a key is right *in principle*, so it does **not** replace this review.

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
| `ActionTab` | data sheet for this op (leave blank for a pattern `DELETE` — it needs no data) |
| `ActionFilter` | *(optional)* value in an `Action` column to select rows (e.g. `Insert`) — see "Data tabs" below |
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

### Data tabs — keep them lean

A data tab holds **one row per record, with a column for each `CELL` field only**.
Everything else is supplied by the tool, so **don't put it in the tab**:

- `PARAM` / `CONST` / `ECHO` / `NULL` columns are generated — omit them from the data.
- Extra columns the operation doesn't use (notes, etc.) are **ignored**.
- A **pattern `DELETE`** is fully defined by its `FromPredicate` — it reads **no** data tab (leave `ActionTab` blank).
- The **`Action` column is optional** and is read **only** when the op sets `ActionFilter` — use it solely to keep non-actionable rows (e.g. `Already in prod`) in the same sheet and have them skipped. On a single-purpose tab, omit it. *(The engine prints a `WARNING` if it finds an `Action` column with no `ActionFilter` — meaning it's being ignored — so the workbook can't quietly mislead.)*

The two samples show both conventions: `tmgGlobalCodes` keeps an `Action` column **and** a reference row to demonstrate the filter; `tmdHTSAdditional` is lean — no `Action` column, and the `DELETE` has no tab.

The generated SQL is **deterministic** — same workbook in, byte-identical SQL out (the backup table name is keyed on Feature + StoryId, not a generation date).

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
python gen_dba_script.py --workbook samples/STD_tmgGlobalCodes_5463147.xlsx --out-dir samples/out --confirm-operations
python gen_dba_script.py --workbook samples/STD_tmdHTSAdditional_5462916.xlsx --out-dir samples/out --confirm-operations
```

---

## 7. What it does *not* decide for you

- **The match key** (see warning above) — explicit and reviewable, not inferred.
- **Pattern deletes** (cleaning up malformed rows) need an explicit `FromPredicate`.
- **Business rules** (`NeverDelete`, "PartnerID never hardcoded") are config flags you set.
- Garbage data in → garbage SQL out. The standard removes structural ambiguity, not domain correctness.

---

## 8. Provenance

Validated against already-shipped, human-reviewed deliverables in `gtm-legacy_gtm-sql`:

- **Generator** — the `tmdHTSAdditional` deploy is **data-identical** to the QA-validated
  `V26.2.0713`, and the `tmgGlobalCodes` deploy is **semantically equivalent** to the merged
  `V26.2.0714`.
- **Adapter** (`build_workbook.py`) — rebuilding the `5441030` and `5346291` workbooks from
  specs yields deploy + verify SQL **byte-identical** to separately-validated output; the
  multi-op `5475122` remediation reproduces the correct 279/16/16/1092 operations with the
  keyed DELETE and both idempotency guards.

See [DESIGN.md](DESIGN.md) for the full background, the unifying model, and limits.
