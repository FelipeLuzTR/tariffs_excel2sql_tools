/* ============================================================
   US 5463147 -- 232_Metals_CSMS68855869 -- VERIFICATION (read-only)
   Run AFTER the deploy script. Every roll-up column should read 'PASS'.
   Counts are payload-scoped so pre-existing rows never inflate them.
============================================================ */
SET NOCOUNT ON;
DECLARE @PartnerID INT = (SELECT TOP 1 PartnerID FROM dbo.tmfDefaults WITH (NOLOCK));

/* ---- Backup exists (AC-1/AC-2) ---- */
DECLARE @v_BackupRows INT = NULL;
IF OBJECT_ID(N'[bck].[bck_tmgGlobalCodes_232_Metals_CSMS68855869_5463147]','U') IS NOT NULL
BEGIN
    DECLARE @v_bsql nvarchar(max) = N'SELECT @c = COUNT(*) FROM [bck].[bck_tmgGlobalCodes_232_Metals_CSMS68855869_5463147] WITH (NOLOCK)';
    EXEC sys.sp_executesql @v_bsql, N'@c INT OUTPUT', @c = @v_BackupRows OUTPUT;
END
SELECT [AC] = 'Backup', [BackupTable] = N'[bck].[bck_tmgGlobalCodes_232_Metals_CSMS68855869_5463147]',
       [Exists] = CASE WHEN OBJECT_ID(N'[bck].[bck_tmgGlobalCodes_232_Metals_CSMS68855869_5463147]','U') IS NOT NULL THEN 1 ELSE 0 END, [RowCount] = @v_BackupRows;

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
     [Backup exists] = CASE WHEN OBJECT_ID(N'[bck].[bck_tmgGlobalCodes_232_Metals_CSMS68855869_5463147]','U') IS NOT NULL THEN 'PASS' ELSE 'FAIL' END
    ,[Op1 per FieldName] = CASE WHEN NOT EXISTS (SELECT 1 FROM @v_exp1 e WHERE e.[expected] <> (SELECT COUNT(*) FROM dbo.tmgGlobalCodes t WITH (NOLOCK) WHERE EXISTS (SELECT 1 FROM @v_ins1 s WHERE t.[PartnerID] = @PartnerID AND t.[FieldName] = s.[FieldName] AND t.[Code] = s.[Code]) AND t.[FieldName] = e.[grp])) THEN 'PASS' ELSE 'FAIL' END
    ,[Op1 no-dup] = CASE WHEN NOT EXISTS (SELECT 1 FROM dbo.tmgGlobalCodes t WITH (NOLOCK) WHERE EXISTS (SELECT 1 FROM @v_ins1 s WHERE t.[PartnerID] = @PartnerID AND t.[FieldName] = s.[FieldName] AND t.[Code] = s.[Code]) GROUP BY t.[FieldName], t.[Code] HAVING COUNT(*)>1) THEN 'PASS' ELSE 'FAIL' END;
