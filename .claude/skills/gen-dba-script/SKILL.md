---
name: gen-dba-script
description: This skill should be used when the user has a completed standardized DBA workbook (an .xlsx with _Meta / _Columns / _Operations control sheets) and asks to "generate the deploy script", "run the DBA generator", "produce the SQL from this workbook", or "generate the deploy/verify/dev-test SQL". It runs the committed, deterministic generator that emits the deploy, verify, and QA dev-test SQL from one workbook — behind a human operations-approval gate.
version: 0.2.0
---

# Generate DBA scripts from a standardized workbook

Turn one standardized Excel workbook into three SQL artifacts using the committed,
deterministic generator. Never hand-write this SQL — always run the generator so the
output is consistent, idempotent, and matches the validated patterns.

## When to use

Use when a workbook already exists in the standardized format — control sheets
`_Meta`, `_Columns`, `_Operations` plus action data tabs. To *produce* a workbook first:
from a regulatory bulletin, use the `csms-to-dba-script` skill; from a BA's raw delta
attachment, use the committed adapter `dba-script-generator/build_workbook.py` with a
per-table profile + a per-story spec (README §2). The full workbook contract is in
`dba-script-generator/README.md` (§5) and `dba-script-generator/DESIGN.md`.

## Two trust tiers (read this first)

A workbook has two tiers, scrutinized differently:

- **Action-tab DATA** — the rows. BA/SME-sourced; assumed authoritative.
- **`_Meta` / `_Columns` / `_Operations`** — the **engineering layer**: match keys, idempotency guards, operation order. When this was **AI-authored** (built from a bulletin, an adapter, or hand-rolled by an agent), it **MUST be human-reviewed before any SQL is written.** The **match key and guards are the highest-risk decisions**, and the dev-test does **not** catch a key that is "right for this dataset but wrong in principle" (the 4-vs-6-column trap). So the dev-test validates the data, not the interpretation.

## How to run — review, THEN confirm

The generator **will not write SQL unless `--confirm-operations` is passed.** Run it in two steps:

**1. Review (no flag → prints the operations, writes nothing):**
```bash
python dba-script-generator/gen_dba_script.py --workbook <path.xlsx> --out-dir <output-dir>
```
This prints the `_Meta` + `_Operations` interpretation (per op: match key, guard, set, idempotency). **Present it to the user/BA and get explicit approval** of the operation interpretation — especially the match keys and guards. Treat any AI-authored `_Operations` as **unapproved** until a human says otherwise.

**2. Generate (only after explicit approval):**
```bash
python dba-script-generator/gen_dba_script.py --workbook <path.xlsx> --out-dir <output-dir> --confirm-operations
```
It writes three files with production-ready, release-aware names derived from `_Meta`:

- `V<Release>.XXXX__DATA_<Table>_<Feature>_<StoryId>.sql` — deploy (idempotent, backup-first, single transaction; `XXXX` is the Flyway sequence the DBA assigns at integration)
- `VERIFY_<Table>_<Feature>_<StoryId>.sql` — read-only acceptance checks (payload-scoped)
- `DEVTEST_<Table>_<StoryId>_DONOTCOMMIT.sql` — a QA dry-run that applies everything, verifies it, then `ROLLBACK`s

**Never pass `--confirm-operations` without an explicit human OK on the operations.** The flag asserts that review happened; passing it unreviewed defeats the gate.

## After generating

- Surface any `WARNING:` lines the generator prints (e.g. an `Action` column with no `ActionFilter` — it is being ignored).
- Tell the user the run order: **dev-test on QA first** (expect every roll-up column = `PASS`), then the deploy script for real, then the verify script for sign-off.
- Never edit generated SQL by hand. To change the output, edit the workbook and regenerate (generation is deterministic — same workbook in, byte-identical SQL out).
