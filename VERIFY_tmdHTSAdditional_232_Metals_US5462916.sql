/* ============================================================
   US 5462916 -- Section 232 Metals (CSMS 68855869) -- VERIFICATION
   ACCEPTANCE-CRITERIA QUERIES (read-only).  Run AFTER the deploy script.
   All SELECTs use WITH (NOLOCK).  Every roll-up column should read 'PASS'.

   NOTE on AC-3 / AC-4 literal wording vs. data:
     * AC-3 literal "count of 99038212/232 = 0" only holds on the FIRST run
       BEFORE the insert step. Post-script the heading legitimately holds 344
       correct rows, so AC-3 is verified here as "0 BROKEN-pattern rows remain".
     * AC-4 literal blanket "TariffType=232 AND EndEffDate='2026-06-07 23:59:59'"
       returns 179 updated + 50 inserted rows that legitimately share that end
       date. AC-4 is therefore verified here scoped to the 179 update keys.
============================================================ */

SET NOCOUNT ON;
DECLARE @BackupHTSTableName SYSNAME = N'[bck].[bck_tmdHTSAdditional_Backup_US_5462916]';

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

/* ---- AC-4 : the 179 update-target rows now carry EndEffDate '2026-06-07 23:59:59' ---- */
DECLARE @EndKeys TABLE (HTSNum varchar(20), Chapter99 varchar(20), CountryofOrigin varchar(10), TariffType varchar(20));
    INSERT INTO @EndKeys (HTSNum, Chapter99, CountryofOrigin, TariffType) VALUES
        (N'87082921', N'99038203', N'', N'232'),
        (N'87082921', N'99038201', N'', N'232'),
        (N'84079010', N'99038205', N'', N'232'),
        (N'84151060', N'99038205', N'', N'232'),
        (N'84151090', N'99038205', N'', N'232'),
        (N'84158101', N'99038205', N'', N'232'),
        (N'84158201', N'99038205', N'', N'232'),
        (N'84159080', N'99038205', N'', N'232'),
        (N'8415908010', N'99038205', N'', N'232'),
        (N'8415908020', N'99038205', N'', N'232'),
        (N'8415908045', N'99038205', N'', N'232'),
        (N'8415908085', N'99038205', N'', N'232'),
        (N'84198150', N'99038205', N'', N'232'),
        (N'84271040', N'99038205', N'', N'232'),
        (N'84271080', N'99038205', N'', N'232'),
        (N'84272040', N'99038205', N'', N'232'),
        (N'84272080', N'99038205', N'', N'232'),
        (N'84279000', N'99038205', N'', N'232'),
        (N'84291100', N'99038205', N'', N'232'),
        (N'84291900', N'99038205', N'', N'232'),
        (N'84292000', N'99038205', N'', N'232'),
        (N'84293000', N'99038205', N'', N'232'),
        (N'84294000', N'99038205', N'', N'232'),
        (N'84295110', N'99038205', N'', N'232'),
        (N'84295150', N'99038205', N'', N'232'),
        (N'84295210', N'99038205', N'', N'232'),
        (N'84295250', N'99038205', N'', N'232'),
        (N'84295910', N'99038205', N'', N'232'),
        (N'84295950', N'99038205', N'', N'232'),
        (N'84312000', N'99038205', N'', N'232'),
        (N'84314200', N'99038205', N'', N'232'),
        (N'84314990', N'99038205', N'', N'232'),
        (N'84321000', N'99038205', N'', N'232'),
        (N'84329000', N'99038205', N'', N'232'),
        (N'84332000', N'99038205', N'', N'232'),
        (N'84335100', N'99038205', N'', N'232'),
        (N'84335900', N'99038205', N'', N'232'),
        (N'84339050', N'99038205', N'', N'232'),
        (N'84798955', N'99038205', N'', N'232'),
        (N'84798965', N'99038205', N'', N'232'),
        (N'85162900', N'99038205', N'', N'232'),
        (N'87011001', N'99038205', N'', N'232'),
        (N'87013010', N'99038205', N'', N'232'),
        (N'87013050', N'99038205', N'', N'232'),
        (N'87019110', N'99038205', N'', N'232'),
        (N'87019150', N'99038205', N'', N'232'),
        (N'87019210', N'99038205', N'', N'232'),
        (N'87019250', N'99038205', N'', N'232'),
        (N'87019310', N'99038205', N'', N'232'),
        (N'87019350', N'99038205', N'', N'232'),
        (N'87019410', N'99038205', N'', N'232'),
        (N'87019450', N'99038205', N'', N'232'),
        (N'87019510', N'99038205', N'', N'232'),
        (N'87019550', N'99038205', N'', N'232'),
        (N'87032101', N'99038205', N'', N'232'),
        (N'87051000', N'99038205', N'', N'232'),
        (N'87052000', N'99038205', N'', N'232'),
        (N'87060030', N'99038205', N'', N'232'),
        (N'87082921', N'99038205', N'', N'232'),
        (N'87084030', N'99038205', N'', N'232'),
        (N'87084060', N'99038205', N'', N'232'),
        (N'87089210', N'99038205', N'', N'232'),
        (N'87089260', N'99038205', N'', N'232'),
        (N'87089315', N'99038205', N'', N'232'),
        (N'87089330', N'99038205', N'', N'232'),
        (N'87089923', N'99038205', N'', N'232'),
        (N'87168010', N'99038205', N'', N'232'),
        (N'87169010', N'99038205', N'', N'232'),
        (N'8415908010', N'99038206', N'', N'232'),
        (N'8415908020', N'99038206', N'', N'232'),
        (N'8415908045', N'99038206', N'', N'232'),
        (N'8415908085', N'99038206', N'', N'232'),
        (N'87082921', N'99038206', N'', N'232'),
        (N'84079010', N'99038209', N'', N'232'),
        (N'84151060', N'99038209', N'', N'232'),
        (N'84151090', N'99038209', N'', N'232'),
        (N'84158101', N'99038209', N'', N'232'),
        (N'84158201', N'99038209', N'', N'232'),
        (N'84159080', N'99038209', N'', N'232'),
        (N'8415908010', N'99038209', N'', N'232'),
        (N'8415908020', N'99038209', N'', N'232'),
        (N'8415908045', N'99038209', N'', N'232'),
        (N'8415908085', N'99038209', N'', N'232'),
        (N'84198150', N'99038209', N'', N'232'),
        (N'84321000', N'99038209', N'', N'232'),
        (N'84329000', N'99038209', N'', N'232'),
        (N'84332000', N'99038209', N'', N'232'),
        (N'84335100', N'99038209', N'', N'232'),
        (N'84335900', N'99038209', N'', N'232'),
        (N'84339050', N'99038209', N'', N'232'),
        (N'84798955', N'99038209', N'', N'232'),
        (N'84798965', N'99038209', N'', N'232'),
        (N'85162900', N'99038209', N'', N'232'),
        (N'87013010', N'99038209', N'', N'232'),
        (N'87019110', N'99038209', N'', N'232'),
        (N'87019210', N'99038209', N'', N'232'),
        (N'87019310', N'99038209', N'', N'232'),
        (N'87019410', N'99038209', N'', N'232'),
        (N'87019510', N'99038209', N'', N'232'),
        (N'87032101', N'99038209', N'', N'232'),
        (N'87060030', N'99038209', N'', N'232'),
        (N'87082921', N'99038209', N'', N'232'),
        (N'87084030', N'99038209', N'', N'232'),
        (N'87084060', N'99038209', N'', N'232'),
        (N'87089210', N'99038209', N'', N'232'),
        (N'87089260', N'99038209', N'', N'232'),
        (N'87089315', N'99038209', N'', N'232'),
        (N'87089330', N'99038209', N'', N'232'),
        (N'87089923', N'99038209', N'', N'232'),
        (N'87168010', N'99038209', N'', N'232'),
        (N'87169010', N'99038209', N'', N'232'),
        (N'84079010', N'99038215', N'RU', N'232'),
        (N'84151060', N'99038215', N'RU', N'232'),
        (N'84151090', N'99038215', N'RU', N'232'),
        (N'84158101', N'99038215', N'RU', N'232'),
        (N'84158201', N'99038215', N'RU', N'232'),
        (N'84159080', N'99038215', N'RU', N'232'),
        (N'84198150', N'99038215', N'RU', N'232'),
        (N'84321000', N'99038215', N'RU', N'232'),
        (N'84329000', N'99038215', N'RU', N'232'),
        (N'84332000', N'99038215', N'RU', N'232'),
        (N'84335100', N'99038215', N'RU', N'232'),
        (N'84335900', N'99038215', N'RU', N'232'),
        (N'84339050', N'99038215', N'RU', N'232'),
        (N'84798955', N'99038215', N'RU', N'232'),
        (N'84798965', N'99038215', N'RU', N'232'),
        (N'85162900', N'99038215', N'RU', N'232'),
        (N'87013010', N'99038215', N'RU', N'232'),
        (N'87019110', N'99038215', N'RU', N'232'),
        (N'87019210', N'99038215', N'RU', N'232'),
        (N'87019310', N'99038215', N'RU', N'232'),
        (N'87019410', N'99038215', N'RU', N'232'),
        (N'87019510', N'99038215', N'RU', N'232'),
        (N'87032101', N'99038215', N'RU', N'232'),
        (N'87060030', N'99038215', N'RU', N'232'),
        (N'87082921', N'99038215', N'RU', N'232'),
        (N'87084030', N'99038215', N'RU', N'232'),
        (N'87084060', N'99038215', N'RU', N'232'),
        (N'87089210', N'99038215', N'RU', N'232'),
        (N'87089260', N'99038215', N'RU', N'232'),
        (N'87089315', N'99038215', N'RU', N'232'),
        (N'87089330', N'99038215', N'RU', N'232'),
        (N'87089923', N'99038215', N'RU', N'232'),
        (N'87168010', N'99038215', N'RU', N'232'),
        (N'87169010', N'99038215', N'RU', N'232'),
        (N'84079010', N'99038216', N'RU', N'232'),
        (N'84151060', N'99038216', N'RU', N'232'),
        (N'84151090', N'99038216', N'RU', N'232'),
        (N'84158101', N'99038216', N'RU', N'232'),
        (N'84158201', N'99038216', N'RU', N'232'),
        (N'84159080', N'99038216', N'RU', N'232'),
        (N'84198150', N'99038216', N'RU', N'232'),
        (N'84321000', N'99038216', N'RU', N'232'),
        (N'84329000', N'99038216', N'RU', N'232'),
        (N'84332000', N'99038216', N'RU', N'232'),
        (N'84335100', N'99038216', N'RU', N'232'),
        (N'84335900', N'99038216', N'RU', N'232'),
        (N'84339050', N'99038216', N'RU', N'232'),
        (N'84798955', N'99038216', N'RU', N'232'),
        (N'84798965', N'99038216', N'RU', N'232'),
        (N'85162900', N'99038216', N'RU', N'232'),
        (N'87013010', N'99038216', N'RU', N'232'),
        (N'87019110', N'99038216', N'RU', N'232'),
        (N'87019210', N'99038216', N'RU', N'232'),
        (N'87019310', N'99038216', N'RU', N'232'),
        (N'87019410', N'99038216', N'RU', N'232'),
        (N'87019510', N'99038216', N'RU', N'232'),
        (N'87032101', N'99038216', N'RU', N'232'),
        (N'87060030', N'99038216', N'RU', N'232'),
        (N'87082921', N'99038216', N'RU', N'232'),
        (N'87084030', N'99038216', N'RU', N'232'),
        (N'87084060', N'99038216', N'RU', N'232'),
        (N'87089210', N'99038216', N'RU', N'232'),
        (N'87089260', N'99038216', N'RU', N'232'),
        (N'87089315', N'99038216', N'RU', N'232'),
        (N'87089330', N'99038216', N'RU', N'232'),
        (N'87089923', N'99038216', N'RU', N'232'),
        (N'87168010', N'99038216', N'RU', N'232'),
        (N'87169010', N'99038216', N'RU', N'232');
SELECT [AC] = 'AC-4 endeff', [Expected] = (SELECT COUNT(*) FROM @EndKeys),
       [MatchedWithNewDate] = COUNT(*)
FROM dbo.tmdHTSAdditional t WITH (NOLOCK)
JOIN @EndKeys k ON t.HTSNum=k.HTSNum AND t.Chapter99=k.Chapter99 AND t.TariffType=k.TariffType
   AND ISNULL(t.CountryofOrigin,'')=ISNULL(k.CountryofOrigin,'')
WHERE t.EndEffDate = CAST(N'2026-06-07 23:59:59' AS datetime);

/* ---- AC-5 : the 17 update-target rows now carry StartEffDate '2026-06-08 00:00:00' ---- */
DECLARE @StartKeys TABLE (HTSNum varchar(20), Chapter99 varchar(20), CountryofOrigin varchar(10), TariffType varchar(20));
    INSERT INTO @StartKeys (HTSNum, Chapter99, CountryofOrigin, TariffType) VALUES
        (N'85169050', N'99038201', N'', N'232'),
        (N'85169050', N'99038203', N'', N'232'),
        (N'84069040', N'99038209', N'', N'232'),
        (N'84109000', N'99038209', N'', N'232'),
        (N'84431600', N'99038209', N'', N'232'),
        (N'85023100', N'99038209', N'', N'232'),
        (N'86071903', N'99038209', N'', N'232'),
        (N'84014000', N'99038210', N'', N'232'),
        (N'84179000', N'99038210', N'', N'232'),
        (N'84559040', N'99038210', N'', N'232'),
        (N'84559080', N'99038210', N'', N'232'),
        (N'90139080', N'99038210', N'', N'232'),
        (N'84014000', N'99038211', N'', N'232'),
        (N'84179000', N'99038211', N'', N'232'),
        (N'84559040', N'99038211', N'', N'232'),
        (N'84559080', N'99038211', N'', N'232'),
        (N'90139080', N'99038211', N'', N'232');
SELECT [AC] = 'AC-5 starteff', [Expected] = (SELECT COUNT(*) FROM @StartKeys),
       [MatchedWithNewDate] = COUNT(*)
FROM dbo.tmdHTSAdditional t WITH (NOLOCK)
JOIN @StartKeys k ON t.HTSNum=k.HTSNum AND t.Chapter99=k.Chapter99 AND t.TariffType=k.TariffType
   AND ISNULL(t.CountryofOrigin,'')=ISNULL(k.CountryofOrigin,'')
WHERE t.StartEffDate = CAST(N'2026-06-08 00:00:00' AS datetime);

/* ---- AC-6 : inserted record counts per Chapter 99 (TariffType 232) ---- */
DECLARE @Expected TABLE (Chapter99 varchar(20), ExpectedCount int);
INSERT INTO @Expected (Chapter99, ExpectedCount) VALUES
    ('99038201', 7),
    ('99038203', 7),
    ('99038205', 4),
    ('99038206', 284),
    ('99038207', 39),
    ('99038208', 39),
    ('99038209', 9),
    ('99038210', 44),
    ('99038211', 44),
    ('99038212', 344),
    ('99038215', 3),
    ('99038216', 3),
    ('99038217', 34),
    ('99038220', 28),
    ('99038221', 28),
    ('99038222', 1036);
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
WHERE t.Chapter99 IN ('99038201','99038203','99038205','99038206','99038207','99038208','99038209','99038210','99038211','99038212','99038215','99038216','99038217','99038220','99038221','99038222','99030306')
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
            WHERE t.EndEffDate = CAST(N'2026-06-07 23:59:59' AS datetime)
        ) = (SELECT COUNT(*) FROM @EndKeys) THEN 'PASS' ELSE 'FAIL' END
    ,[AC-5 starteff 17] = CASE WHEN (
            SELECT COUNT(*) FROM dbo.tmdHTSAdditional t WITH (NOLOCK)
            JOIN @StartKeys k ON t.HTSNum=k.HTSNum AND t.Chapter99=k.Chapter99 AND t.TariffType=k.TariffType
               AND ISNULL(t.CountryofOrigin,'')=ISNULL(k.CountryofOrigin,'')
            WHERE t.StartEffDate = CAST(N'2026-06-08 00:00:00' AS datetime)
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
            WHERE t.Chapter99 IN ('99038201','99038203','99038205','99038206','99038207','99038208','99038209','99038210','99038211','99038212','99038215','99038216','99038217','99038220','99038221','99038222','99030306')
            GROUP BY t.HTSNum, t.Chapter99, t.TariffType, ISNULL(t.CountryofOrigin,'')
            HAVING COUNT(*) > 1
        ) THEN 'PASS' ELSE 'FAIL' END;
