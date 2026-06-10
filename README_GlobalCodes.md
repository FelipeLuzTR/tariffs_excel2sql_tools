# tmg_global_codes_gen.py — tmgGlobalCodes SQL Script Generator

Generates idempotent DBA deployment scripts for inserting HTS codes into `dbo.tmgGlobalCodes`.  
Reads records directly from the Excel attachment on the ADO work item — no manual input required.

---

## Companion to generate_sql.py

| Tool | Target Table | Operations | Data Source |
|------|-------------|------------|-------------|
| `generate_sql.py` | `dbo.tmdHTSAdditional` | DELETE + UPDATE + INSERT | Excel spreadsheet (local file) |
| `tmg_global_codes_gen.py` | `dbo.tmgGlobalCodes` | INSERT only (Business Rules v13 §6) | Excel attachment on ADO work item |

---

## Requirements

- Python 3.10+
- `openpyxl` — `pip install openpyxl`
- Azure CLI with active session — `az login`

---

## Quickstart

```cmd
python tmg_global_codes_gen.py ^
  --work-item 5463147 ^
  --target-dir "D:\dev\ogt\gtm-legacy_gtm-sql\Database\Application\26.2\Hotfix" ^
  --series 07 ^
  --dry-run
```

Remove `--dry-run` to write the files.

---

## How It Works

1. Fetches the ADO work item from `dev.azure.com/tr-tax/TaxTrade`
2. Downloads the Excel attachment (e.g. `CSMS68855869_tmgGlobalCodes_Changes.xlsx`)
3. Reads the **INSERTS** sheet, filters rows where `Action = 'Insert'`
4. Auto-detects the release version from the target directory path (e.g. `26.2`)
5. Finds the next available sequence number within the specified series (e.g. `07xx`)
6. Generates three files:
   - **Deploy script** — idempotent MERGE script for QA and production
   - **Verify script** — AC verification queries (run after deploy)
   - **QA Harness** — rollback test harness (`--with-test` flag, never commit)

---

## Arguments

| Argument | Required | Description |
|----------|----------|-------------|
| `--work-item` / `-w` | Yes | ADO work item ID (e.g. `5463147`) |
| `--target-dir` / `-d` | Yes | Directory to write SQL files into |
| `-r FIELDNAME:CODE` | No | Manual record override — repeat for multiple. Skips ADO fetch. |
| `--series` | No | Sequence prefix to stay within (e.g. `07` targets `0701`–`0799`) |
| `--description` | No | Short tag for filename. Auto-derived from work item title if omitted. |
| `--title` | No | Script header title. Defaults to work item title. |
| `--date` | No | Override effective date `YYYY-MM-DD` (default: today) |
| `--with-test` | No | Also generate the QA rollback test harness |
| `--dry-run` | No | Print SQL to stdout without writing files |

---

## Step-by-Step Workflow

**1. Open the ADO work item** and confirm it has an Excel attachment with an INSERTS tab.

**2. Preview with `--dry-run`** to verify records, filename, and sequence number:
```cmd
python tmg_global_codes_gen.py ^
  --work-item <ID> ^
  --target-dir "D:\dev\ogt\gtm-legacy_gtm-sql\Database\Application\26.X\Hotfix" ^
  --series 07 ^
  --dry-run
```

**3. Generate files** by removing `--dry-run`:
```cmd
python tmg_global_codes_gen.py ^
  --work-item <ID> ^
  --target-dir "D:\dev\ogt\gtm-legacy_gtm-sql\Database\Application\26.X\Hotfix" ^
  --series 07
```

**4. Execute the deploy script** against the target database (QA first).

**5. Run the verify script** — every `[Status]` column should read `PASS`.

**6. Commit only the deploy script** to a feature branch and open a PR targeting `develop`.  
   The `VERIFY_*.sql` and `QA_TEST_*.sql` files are for local use only.

---

## Output Files

### Deploy Script — `V{version}.{seq}__DATA_tmgGlobalCodes_{description}_{work-item}.sql`

- Pre-flight checks: `tmfDefaults`, `tmgGlobalCodes`, `[bck]` schema, PartnerID
- Backup to `[bck].[bck_tmgGlobalCodes_{tag}_{date}_{work-item}]` — skipped on re-run
- MERGE on `(PartnerID, FieldName, Code)` — fully idempotent
- PartnerID sourced from `tmfDefaults` — never hardcoded
- All SELECTs use `WITH (NOLOCK)`
- Full TRY/CATCH with transaction ROLLBACK on error

### Verify Script — `VERIFY_tmgGlobalCodes_{description}_{work-item}.sql`

Read-only AC verification queries. Run after the deploy script.  
Every `[Status]` column should read `PASS`.

### QA Rollback Harness — `QA_TEST_tmgGlobalCodes_{description}_{work-item}.sql` ⚠️ DO NOT COMMIT

Runs the real MERGE twice inside a transaction that **always rolls back**:
- Pass 1: apply inserts (expect N rows)
- Pass 2: re-run (expect 0 rows — proves idempotency)
- AC verification against the uncommitted state
- Post-rollback proof that nothing persisted

---

## Example — US 5463147 (CSMS 68855869)

```cmd
python tmg_global_codes_gen.py ^
  --work-item 5463147 ^
  --target-dir "D:\dev\ogt\gtm-legacy_gtm-sql\Database\Application\26.2\Hotfix" ^
  --series 07
```

**Output:**
```
Fetching work item 5463147 from ADO...
Downloading Excel attachment: CSMS68855869_tmgGlobalCodes_Changes.xlsx
Found 6 record(s) in Excel attachment:
  ABIFTZ-HTS-ALUMINIUM-54RECORD : 37013000
  ABIFTZ-HTS-STEEL-54DERIVATIVE : 8708292120
  ABIFTZ-HTS-STEEL-54DERIVATIVE : 9403200075
  ABIFTZ-HTS-STEEL-54DERIVATIVE : 9403200082
  ABIFTZ-HTS-10PERCENT-DUTYCALC : 99038223
  ABIFTZ-HTS-10PERCENT-DUTYCALC : 9903.82.23

Created deploy  : ...V26.2.0712__DATA_tmgGlobalCodes_232_Metals_CSMS68855869_5463147.sql
Created verify  : ...VERIFY_tmgGlobalCodes_232_Metals_CSMS68855869_5463147.sql
```
