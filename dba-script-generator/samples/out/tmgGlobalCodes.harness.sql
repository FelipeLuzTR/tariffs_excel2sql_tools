/* ============================================================
   *** TEMPORARY -- QA ROLLBACK TEST -- DO NOT COMMIT ***
   US 5463147 -- 232_Metals_CSMS68855869
   Real deploy ops + real verification inside ONE transaction, then ROLLBACK.
   PASS 1 applies; PASS 2 re-runs the same ops (idempotency, expect all 0);
   then AC verification against the uncommitted state. Nothing persists. QA only.
============================================================ */
SET NOCOUNT ON;
SET XACT_ABORT ON;
DECLARE @Msg NVARCHAR(4000) = '"';
DECLARE @CRLF VARCHAR(2) = CHAR(13)+CHAR(10);
DECLARE @EffectiveDate DATETIME = N'2026-06-08 00:00:00';
DECLARE @Op1 INT = 0;
DECLARE @Op1b INT = 0;

IF OBJECT_ID('dbo.tmgGlobalCodes','U') IS NULL
BEGIN
    SET @Msg = @Msg + @CRLF + ' - ERROR: dbo.tmgGlobalCodes does not exist' + CHAR(10) + '"';
    SELECT [Phase]='ERROR', [Msg]=@Msg;
    RETURN;
END
IF OBJECT_ID('dbo.tmfDefaults','U') IS NULL
BEGIN
    SET @Msg = @Msg + @CRLF + ' - ERROR: dbo.tmfDefaults does not exist' + CHAR(10) + '"';
    SELECT [Phase]='ERROR', [Msg]=@Msg;
    RETURN;
END
IF SCHEMA_ID('bck') IS NULL
BEGIN
    SET @Msg = @Msg + @CRLF + ' - ERROR: backup schema does not exist' + CHAR(10) + '"';
    SELECT [Phase]='ERROR', [Msg]=@Msg;
    RETURN;
END
DECLARE @PartnerID INT = (SELECT TOP 1 PartnerID FROM dbo.tmfDefaults WITH (NOLOCK));
IF @PartnerID IS NULL
BEGIN
    SET @Msg = @Msg + @CRLF + ' - ERROR: PartnerID not found in tmfDefaults' + CHAR(10) + '"';
    SELECT [Phase]='ERROR', [Msg]=@Msg;
    RETURN;
END

BEGIN TRY
    BEGIN TRANSACTION;

    /* ===== PASS 1 (apply the real ops) ===== */
    /* ---- Backup (idempotent; never overwritten) ---- */
    IF OBJECT_ID(N'[bck].[bck_tmgGlobalCodes_232_Metals_CSMS68855869_20260610_5463147]','U') IS NULL
    BEGIN
        DECLARE @BackupSQL nvarchar(max) = N'SELECT [PartnerID], [EffDate], [FieldName], [Code], [Decode], [StaticFlag], [DeletedFlag], [KeepDuringRollback] INTO [bck].[bck_tmgGlobalCodes_232_Metals_CSMS68855869_20260610_5463147] FROM dbo.tmgGlobalCodes WITH (NOLOCK)';
        EXEC sys.sp_executesql @BackupSQL;
        SET @Msg = @Msg + @CRLF + ' - Backup created.';
    END
    ELSE SET @Msg = @Msg + @CRLF + ' - Backup already exists. Skipping.';

    /* -- 1. INSERT INSERTS -- */
    DECLARE @Stg1 TABLE (
        [FieldName] varchar(30),
        [Code] nvarchar(36)
    );
    INSERT INTO @Stg1 ([FieldName], [Code]) VALUES
        (N'ABIFTZ-HTS-ALUMINIUM-54RECORD', N'37013000'),
        (N'ABIFTZ-HTS-STEEL-54DERIVATIVE', N'8708292120'),
        (N'ABIFTZ-HTS-STEEL-54DERIVATIVE', N'9403200075'),
        (N'ABIFTZ-HTS-STEEL-54DERIVATIVE', N'9403200082'),
        (N'ABIFTZ-HTS-10PERCENT-DUTYCALC', N'99038223'),
        (N'ABIFTZ-HTS-10PERCENT-DUTYCALC', N'9903.82.23');

    INSERT INTO dbo.tmgGlobalCodes ([PartnerID], [EffDate], [FieldName], [Code], [Decode], [StaticFlag], [DeletedFlag], [KeepDuringRollback])
    SELECT @PartnerID, @EffectiveDate, s.[FieldName], s.[Code], s.[Code], N'Y', N'N', N'N'
    FROM @Stg1 s
    WHERE NOT EXISTS (
        SELECT 1 FROM dbo.tmgGlobalCodes t WITH (NOLOCK)
        WHERE t.[PartnerID] = @PartnerID
          AND t.[FieldName] = s.[FieldName]
          AND t.[Code] = s.[Code]
    );
    SET @Op1 = @@ROWCOUNT;

    SELECT [Phase]='PASS 1 (applied, uncommitted)', [INSERT_1]=@Op1;

    /* ===== PASS 2 (idempotency re-run; expect all 0) ===== */
    /* -- 1. INSERT INSERTS -- */
    INSERT INTO dbo.tmgGlobalCodes ([PartnerID], [EffDate], [FieldName], [Code], [Decode], [StaticFlag], [DeletedFlag], [KeepDuringRollback])
    SELECT @PartnerID, @EffectiveDate, s.[FieldName], s.[Code], s.[Code], N'Y', N'N', N'N'
    FROM @Stg1 s
    WHERE NOT EXISTS (
        SELECT 1 FROM dbo.tmgGlobalCodes t WITH (NOLOCK)
        WHERE t.[PartnerID] = @PartnerID
          AND t.[FieldName] = s.[FieldName]
          AND t.[Code] = s.[Code]
    );
    SET @Op1b = @@ROWCOUNT;

    SELECT [Phase]='PASS 2 (idempotency)', [INSERT_1]=@Op1b,
           [Idempotent]=CASE WHEN @Op1b=0 THEN 'PASS' ELSE 'FAIL' END;

    /* ===== AC VERIFICATION (uncommitted state) ===== */
/* ---- Backup exists (AC-1/AC-2) ---- */
DECLARE @v_BackupRows INT = NULL;
IF OBJECT_ID(N'[bck].[bck_tmgGlobalCodes_232_Metals_CSMS68855869_20260610_5463147]','U') IS NOT NULL
BEGIN
    DECLARE @v_bsql nvarchar(max) = N'SELECT @c = COUNT(*) FROM [bck].[bck_tmgGlobalCodes_232_Metals_CSMS68855869_20260610_5463147] WITH (NOLOCK)';
    EXEC sys.sp_executesql @v_bsql, N'@c INT OUTPUT', @c = @v_BackupRows OUTPUT;
END
SELECT [AC] = 'Backup', [BackupTable] = N'[bck].[bck_tmgGlobalCodes_232_Metals_CSMS68855869_20260610_5463147]',
       [Exists] = CASE WHEN OBJECT_ID(N'[bck].[bck_tmgGlobalCodes_232_Metals_CSMS68855869_20260610_5463147]','U') IS NOT NULL THEN 1 ELSE 0 END, [RowCount] = @v_BackupRows;

/* ---- Op 1 INSERT: payload-scoped count per FieldName ---- */
DECLARE @v_ins1 TABLE (
    [FieldName] varchar(30),
    [Code] nvarchar(36)
);
    INSERT INTO @v_ins1 ([FieldName], [Code]) VALUES
        (N'ABIFTZ-HTS-10PERCENT-DUTYCALC', N'9903.82.23'),
        (N'ABIFTZ-HTS-10PERCENT-DUTYCALC', N'99038223'),
        (N'ABIFTZ-HTS-ALUMINIUM-54RECORD', N'37013000'),
        (N'ABIFTZ-HTS-STEEL-54DERIVATIVE', N'8708292120'),
        (N'ABIFTZ-HTS-STEEL-54DERIVATIVE', N'9403200075'),
        (N'ABIFTZ-HTS-STEEL-54DERIVATIVE', N'9403200082');
DECLARE @v_exp1 TABLE ([grp] nvarchar(60), [expected] int);
INSERT INTO @v_exp1 VALUES
    ('ABIFTZ-HTS-10PERCENT-DUTYCALC', 2),
    ('ABIFTZ-HTS-ALUMINIUM-54RECORD', 1),
    ('ABIFTZ-HTS-STEEL-54DERIVATIVE', 3);
SELECT [AC]='Op1 per FieldName', e.[grp], e.[expected],
       [present] = (SELECT COUNT(*) FROM dbo.tmgGlobalCodes t WITH (NOLOCK) WHERE EXISTS (SELECT 1 FROM @v_ins1 s WHERE t.[PartnerID] = @PartnerID AND t.[FieldName] = s.[FieldName] AND t.[Code] = s.[Code]) AND t.[FieldName] = e.[grp]),
       [Status] = CASE WHEN e.[expected] = (SELECT COUNT(*) FROM dbo.tmgGlobalCodes t WITH (NOLOCK) WHERE EXISTS (SELECT 1 FROM @v_ins1 s WHERE t.[PartnerID] = @PartnerID AND t.[FieldName] = s.[FieldName] AND t.[Code] = s.[Code]) AND t.[FieldName] = e.[grp]) THEN 'PASS' ELSE 'FAIL' END
FROM @v_exp1 e ORDER BY e.[grp];

/* ---- Op 1 INSERT: no duplicates on key (EXPECTED 0) ---- */
SELECT [AC]='Op1 no-dup', t.[FieldName], t.[Code], [n]=COUNT(*)
FROM dbo.tmgGlobalCodes t WITH (NOLOCK)
WHERE EXISTS (SELECT 1 FROM @v_ins1 s WHERE t.[PartnerID] = @PartnerID AND t.[FieldName] = s.[FieldName] AND t.[Code] = s.[Code])
GROUP BY t.[FieldName], t.[Code] HAVING COUNT(*) > 1;

/* ---- Sign-off roll-up (every column should read PASS) ---- */
SELECT
     [Backup exists] = CASE WHEN OBJECT_ID(N'[bck].[bck_tmgGlobalCodes_232_Metals_CSMS68855869_20260610_5463147]','U') IS NOT NULL THEN 'PASS' ELSE 'FAIL' END
    ,[Op1 per FieldName] = CASE WHEN NOT EXISTS (SELECT 1 FROM @v_exp1 e WHERE e.[expected] <> (SELECT COUNT(*) FROM dbo.tmgGlobalCodes t WITH (NOLOCK) WHERE EXISTS (SELECT 1 FROM @v_ins1 s WHERE t.[PartnerID] = @PartnerID AND t.[FieldName] = s.[FieldName] AND t.[Code] = s.[Code]) AND t.[FieldName] = e.[grp])) THEN 'PASS' ELSE 'FAIL' END
    ,[Op1 no-dup] = CASE WHEN NOT EXISTS (SELECT 1 FROM dbo.tmgGlobalCodes t WITH (NOLOCK) WHERE EXISTS (SELECT 1 FROM @v_ins1 s WHERE t.[PartnerID] = @PartnerID AND t.[FieldName] = s.[FieldName] AND t.[Code] = s.[Code]) GROUP BY t.[FieldName], t.[Code] HAVING COUNT(*)>1) THEN 'PASS' ELSE 'FAIL' END;

    IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
    PRINT '*** ROLLED BACK -- nothing persisted. ***';
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
    SELECT [Phase]='ERROR (rolled back)', [ErrLine]=ERROR_LINE(), [ErrMsg]=ERROR_MESSAGE();
END CATCH;
