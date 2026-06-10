#!/usr/bin/env python3
"""
Generic regulatory DBA data-script generator.

Reads ONE standardized Excel workbook (control sheets _Meta / _Columns /
_Operations + action data tabs) and emits THREE artifacts from two shared
builders (single source of truth):

  * deploy  = preamble + guards + BEGIN TRAN + ops_body()            + COMMIT
  * verify  = preamble + verification_body()
  * harness = preamble + guards + BEGIN TRAN + ops_body() [x2 idempotency]
                       + verification_body() + ROLLBACK    (nothing persisted)

Covers both anchor shapes with one engine:
  * tmgGlobalCodes (US 5463147)  -- single idempotent INSERT, partner-scoped
  * tmdHTSAdditional (US 5462916) -- ordered DELETE + 2 UPDATEs + INSERT
See DESIGN.md for the workbook contract.  No database connection required.
"""
import argparse
import datetime as dt
import pandas as pd

CHUNK = 900


# ----------------------------------------------------------------- value helpers
def cell_str(v):
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


def to_dt(v):
    if isinstance(v, pd.Timestamp) or isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    t = cell_str(v)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y"):
        try:
            return dt.datetime.strptime(t, fmt).strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            continue
    raise ValueError(f"unparseable date: {t!r}")


def q(v):
    return "N'" + cell_str(v).replace("'", "''") + "'"


def lit(v, sqltype):
    if "datetime" in sqltype.lower() or sqltype.lower() == "date":
        return f"CAST(N'{to_dt(v)}' AS DATETIME)"
    return q(v)


def values_lines(tuples):
    return ",\n".join("        " + t for t in tuples)


def chunked_insert(table_var, col_sql, tuples):
    out = []
    for i in range(0, len(tuples), CHUNK):
        out.append(f"    INSERT INTO {table_var} ({col_sql}) VALUES\n{values_lines(tuples[i:i+CHUNK])};")
    return "\n".join(out)


# ----------------------------------------------------------------- spec loading
def _kv(df):
    d = {}
    for _, r in df.iterrows():
        k = cell_str(r.iloc[0])
        if k:
            d[k] = r.iloc[1] if len(r) > 1 and not pd.isna(r.iloc[1]) else ""
    return d


def load_all(xlsx):
    xl = pd.ExcelFile(xlsx)
    meta = {k: cell_str(v) for k, v in _kv(pd.read_excel(xl, "_Meta", header=None, dtype=object)).items()}
    columns = [{k: cell_str(v) for k, v in row.items()}
               for _, row in pd.read_excel(xl, "_Columns", dtype=object).fillna("").iterrows()]
    operations = sorted(
        [{k: cell_str(v) for k, v in row.items()}
         for _, row in pd.read_excel(xl, "_Operations", dtype=object).fillna("").iterrows()],
        key=lambda o: int(o["Order"]))
    for op in operations:
        if op.get("ActionTab"):
            df = pd.read_excel(xl, op["ActionTab"], dtype=object)
            df.columns = [str(c).strip() for c in df.columns]
            op["_df"] = df
        else:
            op["_df"] = None
    return meta, columns, operations


def coldef(columns):
    return {c["ColumnName"]: c for c in columns}


def backup_name(meta):
    tbl = meta["TargetTable"].split(".")[-1].strip("[]")
    today = dt.date.today().strftime("%Y%m%d")
    return f"[{meta.get('BackupSchema','bck')}].[bck_{tbl}_{meta['Feature']}_{today}_{meta['StoryId']}]"


def cell_columns(columns):
    return [c["ColumnName"] for c in columns if c["Source"] == "CELL"]


def op_rows(op, columns):
    df = op["_df"]
    if df is None:
        return df
    if op.get("ActionFilter") and "Action" in df.columns:
        return df[df["Action"].map(cell_str) == op["ActionFilter"]]
    return df


# ----------------------------------------------------------------- value sources
def select_expr(col, cdef, s):
    src = cdef[col]["Source"]
    if src == "CELL":
        return f"{s}.[{col}]"
    if src.startswith("CONST:"):
        return q(src[6:])
    if src.startswith("ECHO:"):
        return f"{s}.[{src[5:]}]"
    if src.startswith("PARAM:"):
        return "@" + src[6:]
    if src == "NULL":
        return "NULL"
    return f"{s}.[{col}]"


def key_predicate(col, cdef, s):
    src = cdef[col]["Source"]
    nn = cdef[col].get("NullNormalize", "").upper() == "Y"
    if src.startswith("PARAM:"):
        rhs = "@" + src[6:]
    elif src.startswith("CONST:"):
        rhs = q(src[6:])
    else:
        rhs = f"{s}.[{col}]"
    return f"ISNULL(t.[{col}],'') = ISNULL({rhs},'')" if nn else f"t.[{col}] = {rhs}"


# ----------------------------------------------------------------- op emitters (setup + dml)
def emit_op(op, meta, columns, stg, cnt, declare=True):
    """Returns SQL for one operation. declare=True emits staging DECLARE+populate;
    declare=False emits only the DML (re-using existing staging) -- for the
    harness idempotency re-run."""
    cdef = coldef(columns)
    typ = op["OpType"]

    if typ == "DELETE":
        return (f"    DELETE FROM {meta['TargetTable']}\n    WHERE {op['FromPredicate']};\n"
                f"    SET {cnt} = @@ROWCOUNT;")

    if typ == "INSERT":
        cells = cell_columns(columns)
        rows = op_rows(op, columns)
        all_cols = [c["ColumnName"] for c in columns]
        key = [k.strip() for k in op["MatchKey"].split(",")]
        setup = ""
        if declare:
            tuples = ["(" + ", ".join(lit(r[c], cdef[c]["SqlType"]) for c in cells) + ")"
                      for _, r in rows.iterrows()]
            decl = "    DECLARE " + stg + " TABLE (\n" + ",\n".join(
                f"        [{c}] {cdef[c]['SqlType']}" for c in cells) + "\n    );"
            setup = decl + "\n" + chunked_insert(stg, ", ".join(f"[{c}]" for c in cells), tuples) + "\n\n"
        proj = ", ".join(select_expr(c, cdef, "s") for c in all_cols)
        keypred = "\n          AND ".join(key_predicate(k, cdef, "s") for k in key)
        dml = (f"    INSERT INTO {meta['TargetTable']} ({', '.join('['+c+']' for c in all_cols)})\n"
               f"    SELECT {proj}\n    FROM {stg} s\n    WHERE NOT EXISTS (\n"
               f"        SELECT 1 FROM {meta['TargetTable']} t WITH (NOLOCK)\n        WHERE {keypred}\n    );\n"
               f"    SET {cnt} = @@ROWCOUNT;")
        return setup + dml

    if typ == "UPDATE":
        key = [k.strip() for k in op["MatchKey"].split(",")]
        setcol, _, setsrc = [x.strip() for x in op["SetMap"].partition("<-")]
        new_col = setsrc[5:-1] if setsrc.startswith("CELL(") else None
        set_rhs = f"k.[{new_col}]" if new_col else ("@" + setsrc[6:] if setsrc.startswith("PARAM:") else q(setsrc))
        rows = op_rows(op, columns)
        stg_cols = list(key) + ([new_col] if new_col else [])
        setup = ""
        if declare:
            decls = [f"        [{c}] {cdef[c]['SqlType'] if c in cdef else 'datetime'}" for c in stg_cols]
            tuples = []
            for _, r in rows.iterrows():
                vals = [lit(r[c], cdef[c]["SqlType"]) if c in cdef else q(r[c]) for c in key]
                if new_col:
                    vals.append(lit(r[new_col], cdef.get(setcol, {"SqlType": "datetime"})["SqlType"]))
                tuples.append("(" + ", ".join(vals) + ")")
            decl = "    DECLARE " + stg + " TABLE (\n" + ",\n".join(decls) + "\n    );"
            setup = decl + "\n" + chunked_insert(stg, ", ".join(f"[{c}]" for c in stg_cols), tuples) + "\n\n"
        joinpred = "\n      AND ".join(
            (f"ISNULL(t.[{k}],'') = ISNULL(k.[{k}],'')" if cdef[k].get("NullNormalize", "").upper() == "Y"
             else f"t.[{k}] = k.[{k}]") for k in key)
        where = f"\n    WHERE {op['FromPredicate']}" if op.get("FromPredicate") else ""
        dml = (f"    UPDATE t\n       SET t.[{setcol}] = {set_rhs}\n"
               f"    FROM {meta['TargetTable']} t\n    JOIN {stg} k\n      ON  {joinpred}{where};\n"
               f"    SET {cnt} = @@ROWCOUNT;")
        return setup + dml

    raise ValueError("unknown OpType " + typ)


def ops_body(meta, columns, operations, suffix="", declare=True, include_backup=True):
    """The transaction body: optional backup + each operation, ordered.
    Returns (sql, [counter_var,...], [(label,counter),...])."""
    blocks, counters, summary = [], [], []
    if include_backup and declare:
        bkp = backup_name(meta)
        bcols = ", ".join(f"[{c['ColumnName']}]" for c in columns)
        blocks.append(
            f"""    /* ---- Backup (idempotent; never overwritten) ---- */
    IF OBJECT_ID(N'{bkp}','U') IS NULL
    BEGIN
        DECLARE @BackupSQL nvarchar(max) = N'SELECT {bcols} INTO {bkp} FROM {meta['TargetTable']} WITH (NOLOCK)';
        EXEC sys.sp_executesql @BackupSQL;
        SET @Msg = @Msg + @CRLF + ' - Backup created.';
    END
    ELSE SET @Msg = @Msg + @CRLF + ' - Backup already exists. Skipping.';""")
    for i, op in enumerate(operations, 1):
        cnt = f"@Op{i}{suffix}"
        counters.append(cnt)
        sql = emit_op(op, meta, columns, f"@Stg{i}", cnt, declare=declare)
        blocks.append(f"    /* -- {i}. {op['OpType']} {op.get('ActionTab') or '(pattern)'} -- */\n{sql}")
        summary.append((f"{op['OpType']}_{i}", cnt))
    return "\n\n".join(blocks), counters, summary


# ----------------------------------------------------------------- verification body
def verification_body(meta, columns, operations):
    """AC checks derived from the operations -- payload-scoped, with a PASS/FAIL roll-up.
    Self-contained: declares its own @v_* staging so it never collides with the ops."""
    cdef = coldef(columns)
    bkp = backup_name(meta)
    blocks, rollup = [], []

    # backup existence
    blocks.append(f"""/* ---- Backup exists (AC-1/AC-2) ---- */
DECLARE @v_BackupRows INT = NULL;
IF OBJECT_ID(N'{bkp}','U') IS NOT NULL
BEGIN
    DECLARE @v_bsql nvarchar(max) = N'SELECT @c = COUNT(*) FROM {bkp} WITH (NOLOCK)';
    EXEC sys.sp_executesql @v_bsql, N'@c INT OUTPUT', @c = @v_BackupRows OUTPUT;
END
SELECT [AC] = 'Backup', [BackupTable] = N'{bkp}',
       [Exists] = CASE WHEN OBJECT_ID(N'{bkp}','U') IS NOT NULL THEN 1 ELSE 0 END, [RowCount] = @v_BackupRows;""")
    rollup.append(f"[Backup exists] = CASE WHEN OBJECT_ID(N'{bkp}','U') IS NOT NULL THEN 'PASS' ELSE 'FAIL' END")

    for i, op in enumerate(operations, 1):
        typ = op["OpType"]
        if typ == "DELETE":
            blocks.append(f"""/* ---- Op {i} DELETE: no rows match the delete pattern remain (EXPECTED 0) ---- */
SELECT [AC] = 'Op{i} delete-remaining', [Remaining] = COUNT(*)
FROM {meta['TargetTable']} WITH (NOLOCK) WHERE {op['FromPredicate']};""")
            rollup.append(f"[Op{i} deleted] = CASE WHEN NOT EXISTS (SELECT 1 FROM {meta['TargetTable']} WITH (NOLOCK) "
                          f"WHERE {op['FromPredicate']}) THEN 'PASS' ELSE 'FAIL' END")

        elif typ == "INSERT":
            cells = cell_columns(columns)
            key = [k.strip() for k in op["MatchKey"].split(",")]
            key_cells = [k for k in key if cdef[k]["Source"] == "CELL"]
            rows = op_rows(op, columns)
            ktuples = sorted({"(" + ", ".join(lit(r[c], cdef[c]["SqlType"]) for c in key_cells) + ")"
                              for _, r in rows.iterrows()})
            vk = f"@v_ins{i}"
            decl = "DECLARE " + vk + " TABLE (\n" + ",\n".join(
                f"    [{c}] {cdef[c]['SqlType']}" for c in key_cells) + "\n);"
            pop = chunked_insert(vk, ", ".join(f"[{c}]" for c in key_cells), ktuples)
            keypred = " AND ".join(key_predicate(k, cdef, "s") for k in key)
            grp = op.get("VerifyGroupBy")
            present = (f"(SELECT COUNT(*) FROM {meta['TargetTable']} t WITH (NOLOCK) "
                       f"WHERE EXISTS (SELECT 1 FROM {vk} s WHERE {keypred})")
            if grp:
                exp = rows.groupby(rows[grp].map(cell_str)).size().to_dict()
                exp_rows = ",\n    ".join(f"('{g}', {c})" for g, c in sorted(exp.items()))
                ve = f"@v_exp{i}"
                blocks.append(f"""/* ---- Op {i} INSERT: payload-scoped count per {grp} ---- */
{decl}
{pop}
DECLARE {ve} TABLE ([grp] nvarchar(60), [expected] int);
INSERT INTO {ve} VALUES
    {exp_rows};
SELECT [AC]='Op{i} per {grp}', e.[grp], e.[expected],
       [present] = {present} AND t.[{grp}] = e.[grp]),
       [Status] = CASE WHEN e.[expected] = {present} AND t.[{grp}] = e.[grp]) THEN 'PASS' ELSE 'FAIL' END
FROM {ve} e ORDER BY e.[grp];""")
                rollup.append(f"[Op{i} per {grp}] = CASE WHEN NOT EXISTS (SELECT 1 FROM {ve} e WHERE e.[expected] <> "
                              f"{present} AND t.[{grp}] = e.[grp])) THEN 'PASS' ELSE 'FAIL' END")
            else:
                blocks.append(f"""/* ---- Op {i} INSERT: all payload rows present ---- */
{decl}
{pop}
SELECT [AC]='Op{i} present', [expected]={len(ktuples)},
       [present] = {present});""")
                rollup.append(f"[Op{i} present] = CASE WHEN {present}) = {len(ktuples)} THEN 'PASS' ELSE 'FAIL' END")
            # no-dup on full key (payload-scoped)
            grpkey = ", ".join((f"ISNULL(t.[{k}],'')" if cdef[k].get("NullNormalize", "").upper() == "Y" else f"t.[{k}]")
                               for k in key if cdef[k]["Source"] == "CELL")
            blocks.append(f"""/* ---- Op {i} INSERT: no duplicates on key (EXPECTED 0) ---- */
SELECT [AC]='Op{i} no-dup', {grpkey}, [n]=COUNT(*)
FROM {meta['TargetTable']} t WITH (NOLOCK)
WHERE EXISTS (SELECT 1 FROM {vk} s WHERE {keypred})
GROUP BY {grpkey} HAVING COUNT(*) > 1;""")
            rollup.append(f"[Op{i} no-dup] = CASE WHEN NOT EXISTS (SELECT 1 FROM {meta['TargetTable']} t WITH (NOLOCK) "
                          f"WHERE EXISTS (SELECT 1 FROM {vk} s WHERE {keypred}) GROUP BY {grpkey} HAVING COUNT(*)>1) "
                          f"THEN 'PASS' ELSE 'FAIL' END")

        elif typ == "UPDATE":
            key = [k.strip() for k in op["MatchKey"].split(",")]
            setcol, _, setsrc = [x.strip() for x in op["SetMap"].partition("<-")]
            new_col = setsrc[5:-1] if setsrc.startswith("CELL(") else None
            rows = op_rows(op, columns)
            cols_v = list(key) + ([new_col] if new_col else [])
            tuples = []
            for _, r in rows.iterrows():
                vals = [lit(r[c], cdef[c]["SqlType"]) if c in cdef else q(r[c]) for c in key]
                if new_col:
                    vals.append(lit(r[new_col], cdef.get(setcol, {"SqlType": "datetime"})["SqlType"]))
                tuples.append("(" + ", ".join(vals) + ")")
            vk = f"@v_upd{i}"
            decl = "DECLARE " + vk + " TABLE (\n" + ",\n".join(
                f"    [{c}] {cdef[c]['SqlType'] if c in cdef else 'datetime'}" for c in cols_v) + "\n);"
            pop = chunked_insert(vk, ", ".join(f"[{c}]" for c in cols_v), tuples)
            joinpred = " AND ".join(
                (f"ISNULL(t.[{k}],'') = ISNULL(k.[{k}],'')" if cdef[k].get("NullNormalize", "").upper() == "Y"
                 else f"t.[{k}] = k.[{k}]") for k in key)
            setpred = f"t.[{setcol}] = k.[{new_col}]" if new_col else f"t.[{setcol}] IS NOT NULL"
            blocks.append(f"""/* ---- Op {i} UPDATE: target rows carry the new {setcol} ---- */
{decl}
{pop}
SELECT [AC]='Op{i} update', [expected]=(SELECT COUNT(*) FROM {vk}),
       [matched]=COUNT(*)
FROM {meta['TargetTable']} t WITH (NOLOCK)
JOIN {vk} k ON {joinpred} AND {setpred};""")
            rollup.append(f"[Op{i} update] = CASE WHEN (SELECT COUNT(*) FROM {meta['TargetTable']} t WITH (NOLOCK) "
                          f"JOIN {vk} k ON {joinpred} AND {setpred}) = (SELECT COUNT(*) FROM {vk}) "
                          f"THEN 'PASS' ELSE 'FAIL' END")

    rollup_sql = "SELECT\n     " + "\n    ,".join(rollup) + ";"
    return "\n\n".join(blocks) + "\n\n/* ---- Sign-off roll-up (every column should read PASS) ---- */\n" + rollup_sql


# ----------------------------------------------------------------- preamble / guards
def preamble(meta, columns, operations, extra_counters=()):
    partner = meta.get("PartnerScoped", "").upper() == "Y"
    params = []
    if meta.get("EffectiveDate"):
        params.append(f"DECLARE @EffectiveDate DATETIME = N'{to_dt(meta['EffectiveDate'])}';")
    if meta.get("RetireDate"):
        params.append(f"DECLARE @RetireDate DATETIME = N'{to_dt(meta['RetireDate'])}';")
    decls = ["DECLARE @Msg NVARCHAR(4000) = '\"';", "DECLARE @CRLF VARCHAR(2) = CHAR(13)+CHAR(10);"] + params
    return partner, "\n".join(decls)


def guard_block(meta, columns, operations, diag):
    partner = meta.get("PartnerScoped", "").upper() == "Y"

    def g(cond, msg):
        return ("IF " + cond + "\nBEGIN\n"
                "    SET @Msg = @Msg + @CRLF + ' - ERROR: " + msg + "' + CHAR(10) + '\"';\n"
                "    " + diag + "\n    RETURN;\nEND")
    guards = [g(f"OBJECT_ID('{meta['TargetTable']}','U') IS NULL", meta['TargetTable'] + " does not exist")]
    if partner:
        guards.append(g("OBJECT_ID('dbo.tmfDefaults','U') IS NULL", "dbo.tmfDefaults does not exist"))
    guards.append(g("SCHEMA_ID('" + meta.get('BackupSchema', 'bck') + "') IS NULL", "backup schema does not exist"))
    resolve = ""
    if partner:
        resolve = ("\nDECLARE @PartnerID INT = (" + meta['PartnerSource'] + ");\n"
                   + g("@PartnerID IS NULL", "PartnerID not found in tmfDefaults") + "\n")
    return "\n".join(guards) + resolve


# ----------------------------------------------------------------- builders
def build_deploy(meta, columns, operations):
    _, decls = preamble(meta, columns, operations)
    body, counters, summary = ops_body(meta, columns, operations)
    cdecl = "\n".join(f"DECLARE {c} INT = 0;" for c in counters)
    sel = ", ".join(f"[{lbl}]={c}" for lbl, c in summary)
    bkp = backup_name(meta)
    diag = f"SELECT DatabaseName=DB_NAME(), {sel}, BackupTable=N'{bkp}', Msgs=@Msg;"
    guards = guard_block(meta, columns, operations, diag)
    return f"""/* ============================================================
   US {meta['StoryId']} -- {meta['Feature']}  (target {meta['TargetTable']})
   GENERATED by gen_dba_script.py from a standardized workbook -- do not hand-edit.
   Idempotent: backup-first, set-based, WITH (NOLOCK), single transaction.
============================================================ */
SET NOCOUNT ON;
{decls}
{cdecl}

{guards}
BEGIN TRY
    BEGIN TRANSACTION;

{body}

    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    SET @Msg = @Msg + @CRLF + ' - ERROR (line ' + CAST(ERROR_LINE() AS varchar(10)) + '): ' + ERROR_MESSAGE();
    IF @@TRANCOUNT > 0 BEGIN ROLLBACK TRANSACTION; SET @Msg = @Msg + @CRLF + ' - Rolled back.'; END
END CATCH;

IF LEN(@Msg) > 1 SET @Msg = @Msg + CHAR(10) + '"';
{diag}

/* ============================================================
   QA-ONLY RESTORE (NEVER IN PRODUCTION):
   TRUNCATE TABLE {meta['TargetTable']};
   INSERT INTO {meta['TargetTable']} SELECT * FROM {bkp};
   DROP TABLE {bkp};
============================================================ */
"""


def build_verify(meta, columns, operations):
    partner = meta.get("PartnerScoped", "").upper() == "Y"
    pre = "SET NOCOUNT ON;\n"
    if partner:
        pre += f"DECLARE @PartnerID INT = ({meta['PartnerSource']});\n"
    return f"""/* ============================================================
   US {meta['StoryId']} -- {meta['Feature']} -- VERIFICATION (read-only)
   Run AFTER the deploy script. Every roll-up column should read 'PASS'.
   Counts are payload-scoped so pre-existing rows never inflate them.
============================================================ */
{pre}
{verification_body(meta, columns, operations)}
"""


def build_harness(meta, columns, operations):
    partner = meta.get("PartnerScoped", "").upper() == "Y"
    _, decls = preamble(meta, columns, operations)
    body1, counters1, summary1 = ops_body(meta, columns, operations, suffix="", declare=True)
    body2, counters2, summary2 = ops_body(meta, columns, operations, suffix="b", declare=False, include_backup=False)
    cdecl = "\n".join(f"DECLARE {c} INT = 0;" for c in counters1 + counters2)
    sel1 = ", ".join(f"[{lbl}]={c}" for lbl, c in summary1)
    diag = "SELECT [Phase]='ERROR', [Msg]=@Msg;"
    guards = guard_block(meta, columns, operations, diag)
    idem = " AND ".join(f"{c}=0" for c in counters2)
    return f"""/* ============================================================
   *** TEMPORARY -- QA ROLLBACK TEST -- DO NOT COMMIT ***
   US {meta['StoryId']} -- {meta['Feature']}
   Real deploy ops + real verification inside ONE transaction, then ROLLBACK.
   PASS 1 applies; PASS 2 re-runs the same ops (idempotency, expect all 0);
   then AC verification against the uncommitted state. Nothing persists. QA only.
============================================================ */
SET NOCOUNT ON;
SET XACT_ABORT ON;
{decls}
{cdecl}

{guards}
BEGIN TRY
    BEGIN TRANSACTION;

    /* ===== PASS 1 (apply the real ops) ===== */
{body1}

    SELECT [Phase]='PASS 1 (applied, uncommitted)', {sel1};

    /* ===== PASS 2 (idempotency re-run; expect all 0) ===== */
{body2}

    SELECT [Phase]='PASS 2 (idempotency)', {", ".join(f"[{lbl}]={c}" for lbl, c in summary2)},
           [Idempotent]=CASE WHEN {idem} THEN 'PASS' ELSE 'FAIL' END;

    /* ===== AC VERIFICATION (uncommitted state) ===== */
{verification_body(meta, columns, operations)}

    IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
    PRINT '*** ROLLED BACK -- nothing persisted. ***';
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
    SELECT [Phase]='ERROR (rolled back)', [ErrLine]=ERROR_LINE(), [ErrMsg]=ERROR_MESSAGE();
END CATCH;
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--out", required=True, help="deploy script path")
    ap.add_argument("--out-verify")
    ap.add_argument("--out-test", help="QA rollback harness path")
    args = ap.parse_args()
    meta, columns, operations = load_all(args.workbook)
    open(args.out, "w", encoding="utf-8").write(build_deploy(meta, columns, operations))
    print(f"Deploy : {args.out}")
    if args.out_verify:
        open(args.out_verify, "w", encoding="utf-8").write(build_verify(meta, columns, operations))
        print(f"Verify : {args.out_verify}")
    if args.out_test:
        open(args.out_test, "w", encoding="utf-8").write(build_harness(meta, columns, operations))
        print(f"Harness: {args.out_test}")
    print(f"  target={meta['TargetTable']}  ops={len(operations)}")


if __name__ == "__main__":
    main()
