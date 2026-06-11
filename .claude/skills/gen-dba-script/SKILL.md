---
name: gen-dba-script
description: This skill should be used when the user has a completed standardized DBA workbook (an .xlsx with _Meta / _Columns / _Operations control sheets) and asks to "generate the deploy script", "run the DBA generator", "produce the SQL from this workbook", or "generate the deploy/verify/dev-test SQL". It runs the committed, deterministic generator that emits the deploy, verify, and QA dev-test SQL from one workbook.
version: 0.1.0
---

# Generate DBA scripts from a standardized workbook

Turn one standardized Excel workbook into three SQL artifacts using the committed,
deterministic generator. Never hand-write this SQL — always run the generator so the
output is consistent, idempotent, and matches the validated patterns.

## When to use

Use when a workbook already exists in the standardized format — control sheets
`_Meta`, `_Columns`, `_Operations` plus action data tabs. (To produce a workbook from
a regulatory bulletin first, use the `csms-to-dba-script` skill.) The full workbook
contract is in `dba-script-generator/README.md` (§5) and `dba-script-generator/DESIGN.md`.

## How to run

Run from the repository root (requires `pandas` and `openpyxl`):

```bash
python dba-script-generator/gen_dba_script.py --workbook <path/to/workbook.xlsx> --out-dir <output-dir>
```

It writes three files with production-ready, release-aware names derived from `_Meta`:

- `V<Release>.XXXX__DATA_<Table>_<Feature>_<StoryId>.sql` — deploy (idempotent, backup-first, single transaction; `XXXX` is the Flyway sequence the DBA assigns at integration)
- `VERIFY_<Table>_<Feature>_<StoryId>.sql` — read-only acceptance checks (payload-scoped)
- `DEVTEST_<Table>_<StoryId>_DONOTCOMMIT.sql` — a QA dry-run that applies everything, verifies it, then `ROLLBACK`s

## After generating

- Surface any `WARNING:` lines the generator prints (e.g. an `Action` column with no `ActionFilter` — it is being ignored).
- Tell the user the run order: **dev-test on QA first** (expect every roll-up column = `PASS`), then the deploy script for real, then the verify script for sign-off.
- Never edit generated SQL by hand. To change the output, edit the workbook and regenerate (generation is deterministic — same workbook in, byte-identical SQL out).
