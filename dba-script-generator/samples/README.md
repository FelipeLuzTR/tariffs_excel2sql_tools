# Samples — illustrative fixtures only

**These files are NOT authoritative and do NOT represent deployed changes.**

They exist for two reasons only:
1. **Worked examples** of the standardized workbook format (`STD_*.xlsx`).
2. **Regression proof** that the generator reproduces the validated deliverables (`out/*.sql`).

A file here named after a CSMS or story (e.g. `…CSMS68855869_5462916.sql`) means only that
the generator was once run on that example. It is **not** evidence that the change is
implemented, merged, or deployed.

The authoritative, shipped DBA scripts live in the SQL repository: **`tr/gtm-legacy_gtm-sql`**
(e.g. `Database/Application/<release>/Hotfix/`). To determine whether a change is already
deployed, look there — never infer it from this folder.
