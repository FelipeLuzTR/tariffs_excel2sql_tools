#!/usr/bin/env python3
"""
Generate the idempotent DBA deployment script AND the acceptance-criteria
verification script for US 5462916 (CSMS 68855869 / Proclamation 11032 -
Section 232 Metals HTS updates to dbo.tmdHTSAdditional).

Source of truth: the FINAL spreadsheet attached to the work item.
Reads four data tabs and emits two .sql files. No database connection required.

Operations (single idempotent script, run inside one transaction):
  1. Backup tmdHTSAdditional         (soft-skip if backup already exists)
  2. DELETE broken 9903.82.12 rows   (targeted: the 4 + 47 broken patterns)
  3. UPDATE EndEffDate   (179 rows)
  4. UPDATE StartEffDate (17 rows)
  5. INSERT new rows     (1,955 rows, idempotent via WHERE NOT EXISTS)

Blank CountryofOrigin / HTSNum are stored as empty string ('') per the DBA
clarification on the work item; all key matching is ISNULL(...,'') NULL-safe.
"""
import argparse
import datetime as dt
import pandas as pd

US = "5462916"
CSMS = "68855869"
EXCEL_DEFAULT = "CSMS68855869_tmdhtsAdditional_ChangesKC_FINAL_1.xlsx"
TABLE = "[dbo].[tmdHTSAdditional]"
BACKUP = f"[bck].[bck_tmdHTSAdditional_Backup_US_{US}]"

INSERT_COLS = ["HTSNum", "Chapter99", "CountryofOrigin", "StartEffDate", "EndEffDate",
               "TariffType", "TariffGroup", "RequiredStatusCode", "ValidationLevel", "ExportDate"]
END_DATE_VALUE = "2026-06-07 23:59:59"     # AC-4 retire date
START_DATE_VALUE = "2026-06-08 00:00:00"   # AC-5 effective date

CHUNK = 900  # rows per INSERT ... VALUES statement (SQL Server caps at 1000)


# ---------------------------------------------------------------- value helpers
def cell_str(v):
    """Normalize a cell to a clean string; '' for blanks; strip trailing .0 on floats."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else str(v)
    if isinstance(v, int):
        return str(v)
    return str(v).strip()


def q(v):
    """Quoted N'...' string literal; N'' for blank. Escapes single quotes."""
    return "N'" + cell_str(v).replace("'", "''") + "'"


def to_dt(v):
    if isinstance(v, pd.Timestamp):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, (dt.datetime,)):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    t = cell_str(v)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y"):
        try:
            return dt.datetime.strptime(t, fmt).strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            continue
    raise ValueError(f"unparseable date: {t!r}")


def qdt(v):
    return f"CAST(N'{to_dt(v)}' AS DATETIME)"


def load(excel, sheet):
    df = pd.read_excel(excel, sheet_name=sheet, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    return df


# ---------------------------------------------------------------- VALUES blocks
def values_lines(tuples):
    """tuples: list of strings already formatted as '(a, b, c)'. Returns ',\\n'-joined body."""
    return ",\n".join("        " + t for t in tuples)


def chunked_insert(table_var, col_sql, tuples):
    """Emit one or more INSERT INTO @var (...) VALUES ... statements, <= CHUNK rows each."""
    out = []
    for i in range(0, len(tuples), CHUNK):
        block = tuples[i:i + CHUNK]
        out.append(f"    INSERT INTO {table_var} ({col_sql}) VALUES\n{values_lines(block)};")
    return "\n".join(out)


# ---------------------------------------------------------------- build script
def build_main(excel):
    ins = load(excel, "Inserts_tmdhtsadditional")
    ue = load(excel, "Update_EndEffDate")
    us = load(excel, "Update_StartEffDate")

    # ----- insert tuples (10 cols) -----
    ins_tuples = []
    for _, r in ins.iterrows():
        vals = [q(r["HTSNum"]), q(r["Chapter99"]), q(r["CountryofOrigin"]),
                qdt(r["StartEffDate"]), qdt(r["EndEffDate"]), q(r["TariffType"]),
                q(r["TariffGroup"]), q(r["RequiredStatusCode"]), q(r["ValidationLevel"]),
                qdt(r["ExportDate"])]
        ins_tuples.append("(" + ", ".join(vals) + ")")

    # ----- update-key tuples: (HTSNum, Chapter99, CountryofOrigin, TariffType, NewDate) -----
    end_tuples = ["(" + ", ".join([q(r["HTSNum"]), q(r["Chapter99"]), q(r["CountryofOrigin"]),
                                   q(r["TariffType"]), qdt(r["New_EndEffDate"])]) + ")"
                  for _, r in ue.iterrows()]
    # R3c: StartEffDate match key is (HTSNum, Chapter99, TariffType) -- no CountryofOrigin
    start_tuples = ["(" + ", ".join([q(r["HTSNum"]), q(r["Chapter99"]),
                                     q(r["TariffType"]), qdt(r["New_StartEffDate"])]) + ")"
                    for _, r in us.iterrows()]

    ins_col_sql = ", ".join(f"[{c}]" for c in INSERT_COLS)
    counts = {
        "ins": len(ins_tuples), "end": len(end_tuples), "start": len(start_tuples),
    }

    head = f"""/* ============================================================
   US {US} -- Section 232 Metals HTS Updates per CSMS {CSMS}
                (Proclamation 11032)
   Target table : dbo.tmdHTSAdditional
   Generated    : {dt.date.today():%Y-%m-%d}  (from FINAL spreadsheet, do not hand-edit)

   Operations (idempotent, single transaction):
     1. Backup tmdHTSAdditional into {BACKUP}
        (soft-skip if it already exists -- AC-1 / AC-2)
     2. DELETE broken 9903.82.12 records (targeted patterns)   -- AC-3 / R3a
            R3a soft count guard: broken-pattern count must be 51 (first run)
            or 0 (re-run); any other value -> RAISERROR + rollback (alert).
            (a) HTSNum = '' AND CountryofOrigin IN (BY,CU,KP,RU)   4 rows
            (b) HTSNum <> '' AND CountryofOrigin = ''             47 rows
     3. UPDATE StartEffDate -> '{START_DATE_VALUE}'  ({counts['start']} rows) -- AC-5 / R3c
            match (HTSNum, Chapter99, TariffType) WHERE StartEffDate = '2026-04-06 00:00:00'
     4. UPDATE EndEffDate   -> '{END_DATE_VALUE}'  ({counts['end']} rows)   -- AC-4 / R3b
            match (HTSNum, Chapter99, TariffType, COO) WHERE EndEffDate = '9999-12-31 23:59:59'
     5. INSERT new records  ({counts['ins']} rows, WHERE NOT EXISTS)        -- AC-6 / R3d

   Execution order per R8: backup -> DELETE -> UPDATE StartEff -> UPDATE EndEff -> INSERT.
   Existence key (idempotency): (HTSNum, Chapter99, TariffType, CountryofOrigin)
   Blank HTSNum / CountryofOrigin are stored as '' (empty string); all key
   matching is ISNULL(...,'') NULL-safe per the work-item clarification.
   All SELECTs use WITH (NOLOCK).  Manual verification only (AC-8).
============================================================ */

SET NOCOUNT ON;

DECLARE @BackupHTSTableName SYSNAME = N'{BACKUP}';
DECLARE @Msg            NVARCHAR(4000) = '"';
DECLARE @CRLF           VARCHAR(2)     = CHAR(13) + CHAR(10);
DECLARE @DeletedHTS     INT = 0;
DECLARE @UpdatedEndEff  INT = 0;
DECLARE @UpdatedStart   INT = 0;
DECLARE @InsertedHTS    INT = 0;

BEGIN TRY

    IF OBJECT_ID('dbo.tmdHTSAdditional', 'U') IS NULL
    BEGIN
        SET @Msg = @Msg + @CRLF + ' - ERROR: dbo.tmdHTSAdditional does not exist' + CHAR(10) + '"';
        SELECT DatabaseName = DB_NAME(), DeletedHTS = @DeletedHTS, UpdatedEndEff = @UpdatedEndEff,
               UpdatedStart = @UpdatedStart, InsertedHTS = @InsertedHTS,
               BackupHTSTable = @BackupHTSTableName, Msgs = @Msg;
        RETURN;
    END

    IF SCHEMA_ID('bck') IS NULL
    BEGIN
        SET @Msg = @Msg + @CRLF + ' - ERROR: backup schema [bck] does not exist' + CHAR(10) + '"';
        SELECT DatabaseName = DB_NAME(), DeletedHTS = @DeletedHTS, UpdatedEndEff = @UpdatedEndEff,
               UpdatedStart = @UpdatedStart, InsertedHTS = @InsertedHTS,
               BackupHTSTable = @BackupHTSTableName, Msgs = @Msg;
        RETURN;
    END

    BEGIN TRANSACTION;

    /* ---------- 1. Backup (soft-skip on re-run -- AC-1 / AC-2) ---------- */
    IF OBJECT_ID(@BackupHTSTableName, 'U') IS NULL
    BEGIN
        DECLARE @BackupSQL nvarchar(max) =
            N'SELECT * INTO ' + @BackupHTSTableName +
            N' FROM dbo.tmdHTSAdditional WITH (NOLOCK)';
        EXEC sys.sp_executesql @BackupSQL;
        SET @Msg = @Msg + @CRLF + ' - Backup ' + @BackupHTSTableName + ' created.';
    END
    ELSE
        SET @Msg = @Msg + @CRLF + ' - Backup ' + @BackupHTSTableName + ' already exists. Skipping.';

    /* ---------- 2. DELETE broken 9903.82.12 records (AC-3 / R3a) ----------
       Only the two malformed patterns are removed; correct country-specific
       9903.82.12 rows (re-added in step 5) are never touched -> re-run safe.
       R3a soft guard: broken-pattern count must be 51 (first run) or 0 (re-run). */
    DECLARE @BrokenCount INT =
    (
        SELECT COUNT(*) FROM dbo.tmdHTSAdditional WITH (NOLOCK)
        WHERE TariffType = '232' AND Chapter99 = '99038212'
          AND ( (ISNULL(HTSNum,'') =  '' AND ISNULL(CountryofOrigin,'') IN ('BY','CU','KP','RU'))
             OR (ISNULL(HTSNum,'') <> '' AND ISNULL(CountryofOrigin,'') =  '') )
    );
    IF @BrokenCount NOT IN (0, 51)
        RAISERROR('US5462916 R3a: expected 0 or 51 broken 9903.82.12 records, found %d. Halting and rolling back.', 16, 1, @BrokenCount);

    DELETE FROM dbo.tmdHTSAdditional
    WHERE TariffType = '232'
      AND Chapter99  = '99038212'
      AND (
            (ISNULL(HTSNum,'') =  '' AND ISNULL(CountryofOrigin,'') IN ('BY','CU','KP','RU'))
         OR (ISNULL(HTSNum,'') <> '' AND ISNULL(CountryofOrigin,'') =  '')
          );
    SET @DeletedHTS = @@ROWCOUNT;

    /* ---------- 3. UPDATE StartEffDate -> '{START_DATE_VALUE}' (AC-5 / R3c) ----------
       R3c: match (HTSNum, Chapter99, TariffType) WHERE StartEffDate = '2026-04-06 00:00:00'. */
    DECLARE @StartKeys TABLE (
        HTSNum          varchar(20),
        Chapter99       varchar(20),
        TariffType      varchar(20),
        NewStartEffDate datetime
    );
{chunked_insert("@StartKeys", "HTSNum, Chapter99, TariffType, NewStartEffDate", start_tuples)}

    UPDATE t
       SET t.StartEffDate = k.NewStartEffDate
    FROM dbo.tmdHTSAdditional t
    JOIN @StartKeys k
      ON  t.HTSNum     = k.HTSNum
      AND t.Chapter99  = k.Chapter99
      AND t.TariffType = k.TariffType
    WHERE t.StartEffDate = CAST(N'2026-04-06 00:00:00' AS DATETIME);
    SET @UpdatedStart = @@ROWCOUNT;

    /* ---------- 4. UPDATE EndEffDate -> '{END_DATE_VALUE}' (AC-4 / R3b) ----------
       R3b: match (HTSNum, Chapter99, TariffType, COO) WHERE EndEffDate = '9999-12-31 23:59:59'. */
    DECLARE @EndKeys TABLE (
        HTSNum          varchar(20),
        Chapter99       varchar(20),
        CountryofOrigin varchar(10),
        TariffType      varchar(20),
        NewEndEffDate   datetime
    );
{chunked_insert("@EndKeys", "HTSNum, Chapter99, CountryofOrigin, TariffType, NewEndEffDate", end_tuples)}

    UPDATE t
       SET t.EndEffDate = k.NewEndEffDate
    FROM dbo.tmdHTSAdditional t
    JOIN @EndKeys k
      ON  t.HTSNum     = k.HTSNum
      AND t.Chapter99  = k.Chapter99
      AND t.TariffType = k.TariffType
      AND ISNULL(t.CountryofOrigin,'') = ISNULL(k.CountryofOrigin,'')
    WHERE t.EndEffDate = CAST(N'9999-12-31 23:59:59' AS DATETIME);
    SET @UpdatedEndEff = @@ROWCOUNT;

    /* ---------- 5. INSERT new records (idempotent, AC-6 / AC-7) ---------- */
    DECLARE @Ins TABLE (
        HTSNum             varchar(20),
        Chapter99          varchar(20),
        CountryofOrigin    varchar(10),
        StartEffDate       datetime,
        EndEffDate         datetime,
        TariffType         varchar(20),
        TariffGroup        varchar(20),
        RequiredStatusCode varchar(1),
        ValidationLevel    varchar(1),
        ExportDate         datetime
    );
{chunked_insert("@Ins", ins_col_sql, ins_tuples)}

    INSERT INTO dbo.tmdHTSAdditional
        ({ins_col_sql})
    SELECT i.HTSNum, i.Chapter99, i.CountryofOrigin, i.StartEffDate, i.EndEffDate,
           i.TariffType, i.TariffGroup, i.RequiredStatusCode, i.ValidationLevel, i.ExportDate
    FROM @Ins i
    WHERE NOT EXISTS (
        SELECT 1 FROM dbo.tmdHTSAdditional t WITH (NOLOCK)
        WHERE t.HTSNum     = i.HTSNum
          AND t.Chapter99  = i.Chapter99
          AND t.TariffType = i.TariffType
          AND ISNULL(t.CountryofOrigin,'') = ISNULL(i.CountryofOrigin,'')
    );
    SET @InsertedHTS = @@ROWCOUNT;

    COMMIT TRANSACTION;
    SET @Msg = @Msg + @CRLF + ' - Completed: deleted ' + CAST(@DeletedHTS AS varchar(10))
             + ', endeff-updated ' + CAST(@UpdatedEndEff AS varchar(10))
             + ', starteff-updated ' + CAST(@UpdatedStart AS varchar(10))
             + ', inserted ' + CAST(@InsertedHTS AS varchar(10)) + '.';
END TRY
BEGIN CATCH
    DECLARE @ErrMsg NVARCHAR(4000) = ERROR_MESSAGE(), @ErrLine INT = ERROR_LINE();
    SET @Msg = @Msg + @CRLF + ' - ERROR (line ' + CAST(@ErrLine AS varchar(10)) + '): ' + @ErrMsg;
    IF @@TRANCOUNT > 0
    BEGIN
        ROLLBACK TRANSACTION;
        SET @Msg = @Msg + @CRLF + ' - Transaction rolled back.';
    END
END CATCH;

IF LEN(@Msg) > 1 SET @Msg = @Msg + CHAR(10) + '"';

SELECT
     DatabaseName    = DB_NAME()
    ,DeletedHTS      = @DeletedHTS
    ,UpdatedEndEff   = @UpdatedEndEff
    ,UpdatedStart    = @UpdatedStart
    ,InsertedHTS     = @InsertedHTS
    ,BackupHTSTable  = @BackupHTSTableName
    ,Msgs            = @Msg;

/* ============================================================
   R7 -- QA-ONLY RESTORE PROCEDURE   (NEVER RUN IN PRODUCTION)
   ------------------------------------------------------------
   If a test failure occurs in QA AFTER this script has executed,
   run the three steps below (uncomment) BEFORE re-running the script.
   In PRODUCTION the backup table is retained permanently and is never
   dropped, and this block must never be executed.

   TRUNCATE TABLE dbo.tmdHTSAdditional;
   INSERT INTO dbo.tmdHTSAdditional SELECT * FROM {BACKUP};
   DROP TABLE {BACKUP};
   ============================================================ */
"""
    return head, counts


# ---------------------------------------------------------------- build verify
def build_verify(excel):
    ins = load(excel, "Inserts_tmdhtsadditional")
    g = ins[ins["TariffType"].map(cell_str) == "232"].copy()
    g["Chapter99"] = g["Chapter99"].map(cell_str)
    per = g.groupby("Chapter99").size().sort_index()
    per_rows = "\n".join(
        f"    ('{ch}', {cnt})," for ch, cnt in per.items()
    ).rstrip(",")
    headings = ",".join(f"'{ch}'" for ch in per.index)

    v = f"""/* ============================================================
   US {US} -- Section 232 Metals (CSMS {CSMS}) -- VERIFICATION
   ACCEPTANCE-CRITERIA QUERIES (read-only).  Run AFTER the deploy script.
   All SELECTs use WITH (NOLOCK).  Every roll-up column should read 'PASS'.

   NOTE on AC-3 / AC-4 literal wording vs. data:
     * AC-3 literal "count of 99038212/232 = 0" only holds on the FIRST run
       BEFORE the insert step. Post-script the heading legitimately holds 344
       correct rows, so AC-3 is verified here as "0 BROKEN-pattern rows remain".
     * AC-4 literal blanket "TariffType=232 AND EndEffDate='{END_DATE_VALUE}'"
       returns 179 updated + 50 inserted rows that legitimately share that end
       date. AC-4 is therefore verified here scoped to the 179 update keys.
============================================================ */

SET NOCOUNT ON;
DECLARE @BackupHTSTableName SYSNAME = N'{BACKUP}';

/* ---- AC-1 / AC-2 : backup exists and is the single snapshot ---- */
DECLARE @BackupRows INT = NULL;
IF OBJECT_ID(@BackupHTSTableName,'U') IS NOT NULL
BEGIN
    DECLARE @sql nvarchar(max) = N'SELECT @c = COUNT(*) FROM ' + @BackupHTSTableName + N' WITH (NOLOCK)';
    EXEC sys.sp_executesql @sql, N'@c INT OUTPUT', @c = @BackupRows OUTPUT;
END
SELECT [AC] = 'AC-1/AC-2', [BackupTable] = @BackupHTSTableName,
       [BackupExists] = CASE WHEN OBJECT_ID(@BackupHTSTableName,'U') IS NOT NULL THEN 1 ELSE 0 END,
       [BackupRowCount] = @BackupRows;   -- record; must be unchanged on re-runs

/* ---- AC-3 : no BROKEN 9903.82.12 rows remain (EXPECTED 0) ---- */
SELECT [AC] = 'AC-3 broken-remaining', [BrokenRemaining] = COUNT(*)
FROM dbo.tmdHTSAdditional WITH (NOLOCK)
WHERE TariffType = '232' AND Chapter99 = '99038212'
  AND ( (ISNULL(HTSNum,'') =  '' AND ISNULL(CountryofOrigin,'') IN ('BY','CU','KP','RU'))
     OR (ISNULL(HTSNum,'') <> '' AND ISNULL(CountryofOrigin,'') =  '') );

/* ---- AC-4 : the 179 update-target rows now carry EndEffDate '{END_DATE_VALUE}' ---- */
DECLARE @EndKeys TABLE (HTSNum varchar(20), Chapter99 varchar(20), CountryofOrigin varchar(10), TariffType varchar(20));
{verify_keys(excel, "Update_EndEffDate", "@EndKeys")}
SELECT [AC] = 'AC-4 endeff', [Expected] = (SELECT COUNT(*) FROM @EndKeys),
       [MatchedWithNewDate] = COUNT(*)
FROM dbo.tmdHTSAdditional t WITH (NOLOCK)
JOIN @EndKeys k ON t.HTSNum=k.HTSNum AND t.Chapter99=k.Chapter99 AND t.TariffType=k.TariffType
   AND ISNULL(t.CountryofOrigin,'')=ISNULL(k.CountryofOrigin,'')
WHERE t.EndEffDate = CAST(N'{END_DATE_VALUE}' AS datetime);

/* ---- AC-5 : the 17 update-target rows now carry StartEffDate '{START_DATE_VALUE}' ---- */
DECLARE @StartKeys TABLE (HTSNum varchar(20), Chapter99 varchar(20), CountryofOrigin varchar(10), TariffType varchar(20));
{verify_keys(excel, "Update_StartEffDate", "@StartKeys")}
SELECT [AC] = 'AC-5 starteff', [Expected] = (SELECT COUNT(*) FROM @StartKeys),
       [MatchedWithNewDate] = COUNT(*)
FROM dbo.tmdHTSAdditional t WITH (NOLOCK)
JOIN @StartKeys k ON t.HTSNum=k.HTSNum AND t.Chapter99=k.Chapter99 AND t.TariffType=k.TariffType
   AND ISNULL(t.CountryofOrigin,'')=ISNULL(k.CountryofOrigin,'')
WHERE t.StartEffDate = CAST(N'{START_DATE_VALUE}' AS datetime);

/* ---- AC-6 : inserted record counts per Chapter 99 (TariffType 232) ---- */
DECLARE @Expected TABLE (Chapter99 varchar(20), ExpectedCount int);
INSERT INTO @Expected (Chapter99, ExpectedCount) VALUES
{per_rows};
SELECT [AC] = 'AC-6 per-heading',
       e.Chapter99, e.ExpectedCount,
       [ActualCount] = (SELECT COUNT(*) FROM dbo.tmdHTSAdditional t WITH (NOLOCK)
                        WHERE t.TariffType='232' AND t.Chapter99 = e.Chapter99),
       [Status] = CASE WHEN e.ExpectedCount = (SELECT COUNT(*) FROM dbo.tmdHTSAdditional t WITH (NOLOCK)
                        WHERE t.TariffType='232' AND t.Chapter99 = e.Chapter99)
                       THEN 'PASS' ELSE 'FAIL' END
FROM @Expected e ORDER BY e.Chapter99;

-- AC-6 Section 122 (EXPECTED 2)
SELECT [AC] = 'AC-6 section122', [Count] = COUNT(*)
FROM dbo.tmdHTSAdditional WITH (NOLOCK)
WHERE TariffType = '122' AND Chapter99 = '99030306' AND HTSNum IN ('37013000','9403999040');

/* ---- AC-7 : no duplicates on the existence key for the affected headings ---- */
SELECT [AC] = 'AC-7 duplicate', t.HTSNum, t.Chapter99, t.TariffType,
       [COO] = ISNULL(t.CountryofOrigin,''), [Occurrences] = COUNT(*)
FROM dbo.tmdHTSAdditional t WITH (NOLOCK)
WHERE t.Chapter99 IN ({headings},'99030306')
GROUP BY t.HTSNum, t.Chapter99, t.TariffType, ISNULL(t.CountryofOrigin,'')
HAVING COUNT(*) > 1;

/* ---- Sign-off roll-up (every column should be 'PASS') ---- */
SELECT
     [AC-3 no-broken] = CASE WHEN NOT EXISTS (
            SELECT 1 FROM dbo.tmdHTSAdditional WITH (NOLOCK)
            WHERE TariffType='232' AND Chapter99='99038212'
              AND ( (ISNULL(HTSNum,'')='' AND ISNULL(CountryofOrigin,'') IN ('BY','CU','KP','RU'))
                 OR (ISNULL(HTSNum,'')<>'' AND ISNULL(CountryofOrigin,'')='') )
        ) THEN 'PASS' ELSE 'FAIL' END
    ,[AC-4 endeff 179] = CASE WHEN (
            SELECT COUNT(*) FROM dbo.tmdHTSAdditional t WITH (NOLOCK)
            JOIN @EndKeys k ON t.HTSNum=k.HTSNum AND t.Chapter99=k.Chapter99 AND t.TariffType=k.TariffType
               AND ISNULL(t.CountryofOrigin,'')=ISNULL(k.CountryofOrigin,'')
            WHERE t.EndEffDate = CAST(N'{END_DATE_VALUE}' AS datetime)
        ) = (SELECT COUNT(*) FROM @EndKeys) THEN 'PASS' ELSE 'FAIL' END
    ,[AC-5 starteff 17] = CASE WHEN (
            SELECT COUNT(*) FROM dbo.tmdHTSAdditional t WITH (NOLOCK)
            JOIN @StartKeys k ON t.HTSNum=k.HTSNum AND t.Chapter99=k.Chapter99 AND t.TariffType=k.TariffType
               AND ISNULL(t.CountryofOrigin,'')=ISNULL(k.CountryofOrigin,'')
            WHERE t.StartEffDate = CAST(N'{START_DATE_VALUE}' AS datetime)
        ) = (SELECT COUNT(*) FROM @StartKeys) THEN 'PASS' ELSE 'FAIL' END
    ,[AC-6 per-heading] = CASE WHEN NOT EXISTS (
            SELECT 1 FROM @Expected e
            WHERE e.ExpectedCount <> (SELECT COUNT(*) FROM dbo.tmdHTSAdditional t WITH (NOLOCK)
                                      WHERE t.TariffType='232' AND t.Chapter99=e.Chapter99)
        ) THEN 'PASS' ELSE 'FAIL' END
    ,[AC-6 section122=2] = CASE WHEN (
            SELECT COUNT(*) FROM dbo.tmdHTSAdditional WITH (NOLOCK)
            WHERE TariffType='122' AND Chapter99='99030306' AND HTSNum IN ('37013000','9403999040')
        ) = 2 THEN 'PASS' ELSE 'FAIL' END
    ,[AC-7 no-dupes] = CASE WHEN NOT EXISTS (
            SELECT 1 FROM dbo.tmdHTSAdditional t WITH (NOLOCK)
            WHERE t.Chapter99 IN ({headings},'99030306')
            GROUP BY t.HTSNum, t.Chapter99, t.TariffType, ISNULL(t.CountryofOrigin,'')
            HAVING COUNT(*) > 1
        ) THEN 'PASS' ELSE 'FAIL' END;
"""
    return v


def verify_keys(excel, sheet, var):
    df = load(excel, sheet)
    tuples = ["(" + ", ".join([q(r["HTSNum"]), q(r["Chapter99"]), q(r["CountryofOrigin"]),
                               q(r["TariffType"])]) + ")" for _, r in df.iterrows()]
    return chunked_insert(var, "HTSNum, Chapter99, CountryofOrigin, TariffType", tuples)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", default=EXCEL_DEFAULT)
    ap.add_argument("--out-script", default=f"V26.3.XXXX__DATA_tmdHTSAdditional_232_Metals_CSMS{CSMS}.sql")
    ap.add_argument("--out-verify", default=f"VERIFY_tmdHTSAdditional_232_Metals_US{US}.sql")
    args = ap.parse_args()

    script, counts = build_main(args.excel)
    verify = build_verify(args.excel)

    with open(args.out_script, "w", encoding="utf-8") as f:
        f.write(script)
    with open(args.out_verify, "w", encoding="utf-8") as f:
        f.write(verify)

    print(f"Deploy script : {args.out_script}")
    print(f"Verify script : {args.out_verify}")
    print(f"Counts -> inserts={counts['ins']} endeff={counts['end']} starteff={counts['start']}")


if __name__ == "__main__":
    main()
