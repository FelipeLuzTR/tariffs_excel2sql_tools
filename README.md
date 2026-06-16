# tariffs_excel2sql_tools

**Turn a CBP / regulatory tariff change into production-ready, idempotent GTM DBA SQL —
driven by AI skills, behind a human approval gate.**

The headline capability of this repo is a set of **AI skills** that take a regulatory change
(an **ADO user story + its Excel data attachment**) and produce the three SQL
artifacts a DBA story needs for `tmdHTSAdditional` / `tmgGlobalCodes` — a backup-first,
**idempotent deploy** script, a read-only **verification** script, and a transaction+rollback
**QA dev-test** — while stopping for a developer to approve the SQL interpretation before
anything is written.

> The command-line tools that used to be the whole story are still here, but they are now the
> **lower-level / manual** path — see [§ Manual & lower-level tools](#manual--lower-level-tools).

---

## ⭐ Use the AI skill — just describe what you need

You describe what you need in plain language — you don't hand-write the SQL, and nobody
hand-assembles the standardized workbook. (The BA/SME still fills in the **data spreadsheet** by
hand and attaches it to the story; that's the input the AI translates — not something it
replaces.) For example, point it at an ADO user story that has a data attachment:

```
Deliver ADO story 5475122 end to end: read the story + its Excel data attachment, build the
standardized workbook, generate the deploy/verify/dev-test SQL, and dev-test against QA
partner 2046.
```

### What the AI then does

1. **Gathers** the change — reads the ADO user story and its Excel data attachment.
2. **Drafts the standardized workbook** — the operations, match keys, and idempotency guards
   (the engineering interpretation).
3. **Generates the SQL — and stops at the approval gate.** The generator writes **nothing** until
   a developer reviews and approves the operation interpretation.
4. **Hands off three artifacts:** the **deploy** script (idempotent, backup-first, single
   transaction, `NOLOCK`), the **verify** script (payload-scoped acceptance checks), and the
   **dev-test** (applies on QA → verifies → `ROLLBACK`s — nothing persists).

### Who does what

| Role | Owns | Does |
|---|---|---|
| **BA / SME** | regulatory **intent + data** | writes the story (operations, requirements, dates) and attaches the data spreadsheet. Does **not** write SQL or pick match keys. |
| **AI skills** | the **translation** | draft the workbook + generate the SQL — the drudgery a DBA used to do by hand. Never deploy; always stop at the gate. |
| **Developer** | the **engineering review** | reviews the AI's operations (match keys + guards) at the gate, ratifies the judgment calls, approves; owns the dev-test + deploy. |

### The skills (in `.claude/skills/`)

| Skill / tool | Turns… | …into |
|---|---|---|
| **`build_workbook.py`** + profile + spec | a BA's raw Excel attachment from the story | a standardized workbook |
| **`gen-dba-script`** | a standardized workbook | the deploy / verify / dev-test SQL — **behind the operations-approval gate** |

### Safety model (why you can trust it)

- **Nothing deploys automatically.** The generator refuses to write SQL without explicit
  developer approval of the operations (match keys + guards).
- **The dev-test rolls back** — it applies everything inside a transaction and `ROLLBACK`s, so
  you validate against real QA data without persisting anything.
- **Counts are payload-scoped** and every `SELECT` uses `WITH (NOLOCK)`.

📖 **Full how-to, the workbook contract, and the developer gate-review checklist:
[`dba-script-generator/README.md`](dba-script-generator/README.md).**
🗺️ **How this fits the broader CBP→GTM automation strategy: [`PROCESS-MAP.md`](PROCESS-MAP.md).**

---

## Manual & lower-level tools

If you prefer to drive the engine by hand (no AI in the loop), everything above runs from the
command line.

### The generator (recommended manual path)

`dba-script-generator/` is the committed, deterministic engine: **one standardized Excel
workbook → the three SQL artifacts.** It runs behind the same approval gate.

```bash
pip install pandas openpyxl

# review the operations (writes nothing), then generate after approving
python dba-script-generator/gen_dba_script.py --workbook <wb.xlsx> --out-dir out/
python dba-script-generator/gen_dba_script.py --workbook <wb.xlsx> --out-dir out/ --confirm-operations
```

To build a workbook from a BA's raw attachment without the AI, use the adapter
(`build_workbook.py` + a per-table profile + a per-story spec).

➡️ **Everything about the workbook format, the adapter, the gate, and the worked examples lives
in [`dba-script-generator/README.md`](dba-script-generator/README.md).**

### Legacy row-only tools (`main.py` / `tool.js`)

The original utilities that emit **only the bulk `VALUES` rows** — not the backup, idempotency,
transaction, or verification logic. **Superseded by the generator above** (which produces the
complete, idempotent script); kept for reference and one-off row extraction.

| File | Role |
|------|------|
| [`main.py`](main.py) | Excel `InsertData` sheet → SQL `VALUES` tuples for `tmdHTSAdditional` (10-column contract; strips `.0` off numeric HTS codes; preserves datetime time components) |
| [`tool.js`](tool.js) | post-process those rows — `batch` (split into ≤1000-row `INSERT`s) / `index` (add `RowId`) / `merge` |
| [`.command`](.command) | a reminder of the `main.py` invocation |

```bash
python main.py --excel Changes.xlsx --sheet InsertData --table "[dbo].[tmdHTSAdditional]" --out ROWS.sql
```

> ⚠️ These produce a **row scaffold only** — no backup, no `WHERE NOT EXISTS`, no transaction,
> no `VERIFY` script, and blanks become `N''` (not `NULL`). For a production-grade deliverable,
> use the generator / AI skills above. (The pre-restructure README, with the full `main.py`
> value-formatting and caveat detail, is in the git history.)
