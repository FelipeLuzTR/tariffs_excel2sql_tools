#!/usr/bin/env python3
"""Author the sample standardized workbooks under samples/.

  * tmgGlobalCodes (US 5463147) -- fully self-contained (6 rows hardcoded); always built.
  * tmdHTSAdditional (US 5462916) -- needs the work-item source workbook (the 1,955-row
    InsertData + Update/Delete tabs); built only when --hts-source is supplied. A pre-built
    copy is already committed under samples/, so this is only needed to regenerate it.

Usage:
    python make_samples.py                         # rebuild the tmgGlobalCodes sample
    python make_samples.py --hts-source <FINAL.xlsx>   # also rebuild the tmdHTSAdditional sample
"""
import argparse
import os
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
SAMPLES = os.path.join(HERE, "samples")
FEATURE = "232_Metals_CSMS68855869"


def write_book(path, meta, columns, operations, action_tabs):
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        pd.DataFrame(meta).to_excel(w, sheet_name="_Meta", header=False, index=False)
        pd.DataFrame(columns).to_excel(w, sheet_name="_Columns", index=False)
        pd.DataFrame(operations).to_excel(w, sheet_name="_Operations", index=False)
        for name, df in action_tabs.items():
            df.to_excel(w, sheet_name=name, index=False)
    print("wrote", path)


def build_globalcodes():
    meta = [
        ["TargetTable", "dbo.tmgGlobalCodes"], ["StoryId", "5463147"], ["Feature", FEATURE],
        ["Release", "26.2"], ["EffectiveDate", "2026-06-08 00:00:00"], ["BackupSchema", "bck"],
        ["PartnerScoped", "Y"], ["PartnerSource", "SELECT TOP 1 PartnerID FROM dbo.tmfDefaults WITH (NOLOCK)"],
        ["NeverDelete", "Y"],
    ]
    columns = [
        {"ColumnName": "PartnerID", "SqlType": "int", "Source": "PARAM:PartnerID", "NullNormalize": ""},
        {"ColumnName": "EffDate", "SqlType": "datetime", "Source": "PARAM:EffectiveDate", "NullNormalize": ""},
        {"ColumnName": "FieldName", "SqlType": "varchar(30)", "Source": "CELL", "NullNormalize": ""},
        {"ColumnName": "Code", "SqlType": "nvarchar(36)", "Source": "CELL", "NullNormalize": ""},
        {"ColumnName": "Decode", "SqlType": "nvarchar(36)", "Source": "ECHO:Code", "NullNormalize": ""},
        {"ColumnName": "StaticFlag", "SqlType": "char(1)", "Source": "CONST:Y", "NullNormalize": ""},
        {"ColumnName": "DeletedFlag", "SqlType": "char(1)", "Source": "CONST:N", "NullNormalize": ""},
        {"ColumnName": "KeepDuringRollback", "SqlType": "char(1)", "Source": "CONST:N", "NullNormalize": ""},
    ]
    operations = [{
        "Order": 1, "ActionTab": "INSERTS", "ActionFilter": "Insert", "OpType": "INSERT",
        "MatchKey": "PartnerID,FieldName,Code", "FromPredicate": "", "SetMap": "",
        "Idempotency": "NOT_EXISTS", "VerifyGroupBy": "FieldName",
    }]
    rows = [
        ("Insert", "ABIFTZ-HTS-ALUMINIUM-54RECORD", "37013000"),
        ("Insert", "ABIFTZ-HTS-STEEL-54DERIVATIVE", "8708292120"),
        ("Insert", "ABIFTZ-HTS-STEEL-54DERIVATIVE", "9403200075"),
        ("Insert", "ABIFTZ-HTS-STEEL-54DERIVATIVE", "9403200082"),
        ("Insert", "ABIFTZ-HTS-10PERCENT-DUTYCALC", "99038223"),
        ("Insert", "ABIFTZ-HTS-10PERCENT-DUTYCALC", "9903.82.23"),
        # A reference row kept for context but NOT actioned. ActionFilter='Insert'
        # (in _Operations) makes the engine skip it -- this is the only place the
        # Action column does anything, and it demonstrates why it exists.
        ("Already in prod", "ABIFTZ-HTS-STEEL-54PRIMARY", "72081000"),
    ]
    inserts = pd.DataFrame(rows, columns=["Action", "FieldName", "Code"]).astype(str)
    write_book(os.path.join(SAMPLES, "STD_tmgGlobalCodes_5463147.xlsx"),
               meta, columns, operations, {"INSERTS": inserts})


def build_htsadditional(src):
    meta = [
        ["TargetTable", "dbo.tmdHTSAdditional"], ["StoryId", "5462916"], ["Feature", FEATURE],
        ["Release", "26.2"], ["EffectiveDate", "2026-06-08 00:00:00"], ["RetireDate", "2026-06-07 23:59:59"],
        ["BackupSchema", "bck"], ["PartnerScoped", "N"], ["NeverDelete", "N"],
    ]
    nn = {"HTSNum", "CountryofOrigin"}
    coldefs = [
        ("HTSNum", "varchar(20)"), ("Chapter99", "varchar(20)"), ("CountryofOrigin", "varchar(10)"),
        ("StartEffDate", "datetime"), ("EndEffDate", "datetime"), ("TariffType", "varchar(20)"),
        ("TariffGroup", "varchar(20)"), ("RequiredStatusCode", "varchar(1)"),
        ("ValidationLevel", "varchar(1)"), ("ExportDate", "datetime"),
    ]
    columns = [{"ColumnName": n, "SqlType": t, "Source": "CELL",
                "NullNormalize": "Y" if n in nn else ""} for n, t in coldefs]
    delete_pred = ("Chapter99 = '99038212' AND TariffType = '232' AND "
                   "( (ISNULL(HTSNum,'') = '' AND ISNULL(CountryofOrigin,'') IN ('BY','CU','KP','RU')) "
                   "OR (ISNULL(HTSNum,'') <> '' AND ISNULL(CountryofOrigin,'') = '') )")
    operations = [
        # Pattern DELETE -- fully defined by FromPredicate; needs no data tab.
        {"Order": 1, "ActionTab": "", "ActionFilter": "", "OpType": "DELETE",
         "MatchKey": "", "FromPredicate": delete_pred, "SetMap": "", "Idempotency": "PATTERN", "VerifyGroupBy": ""},
        {"Order": 2, "ActionTab": "Update_StartEffDate", "ActionFilter": "", "OpType": "UPDATE",
         "MatchKey": "HTSNum,Chapter99,TariffType",
         "FromPredicate": "t.StartEffDate = CAST(N'2026-04-06 00:00:00' AS DATETIME)",
         "SetMap": "StartEffDate <- CELL(New_StartEffDate)", "Idempotency": "GUARDED", "VerifyGroupBy": ""},
        {"Order": 3, "ActionTab": "Update_EndEffDate", "ActionFilter": "", "OpType": "UPDATE",
         "MatchKey": "HTSNum,Chapter99,CountryofOrigin,TariffType",
         "FromPredicate": "t.EndEffDate = CAST(N'9999-12-31 23:59:59' AS DATETIME)",
         "SetMap": "EndEffDate <- CELL(New_EndEffDate)", "Idempotency": "GUARDED", "VerifyGroupBy": ""},
        # Single-purpose tab -> no Action column / no ActionFilter needed.
        {"Order": 4, "ActionTab": "Inserts_tmdhtsadditional", "ActionFilter": "", "OpType": "INSERT",
         "MatchKey": "HTSNum,Chapter99,TariffType,CountryofOrigin,StartEffDate", "FromPredicate": "",
         "SetMap": "", "Idempotency": "NOT_EXISTS", "VerifyGroupBy": "Chapter99"},
    ]

    def tab(sheet, keep):
        """Read a source tab and keep ONLY the columns this op uses (drops Action/Notes/etc. fat)."""
        df = pd.read_excel(src, sheet_name=sheet, dtype=object)
        df.columns = [str(c).strip() for c in df.columns]
        return df[keep]

    tabs = {
        "Inserts_tmdhtsadditional": tab("Inserts_tmdhtsadditional", [n for n, _ in coldefs]),
        "Update_StartEffDate": tab("Update_StartEffDate", ["HTSNum", "Chapter99", "TariffType", "New_StartEffDate"]),
        "Update_EndEffDate": tab("Update_EndEffDate", ["HTSNum", "Chapter99", "CountryofOrigin", "TariffType", "New_EndEffDate"]),
    }
    write_book(os.path.join(SAMPLES, "STD_tmdHTSAdditional_5462916.xlsx"),
               meta, columns, operations, tabs)


def build_template():
    """A blank, structured starter workbook -- copy it and fill in your own change."""
    meta = [
        ["TargetTable", "dbo.<YourTable>"], ["StoryId", "<ADO story id>"],
        ["Feature", "<feature_slug e.g. 232_Metals_CSMS00000000>"], ["Release", "26.x"],
        ["EffectiveDate", "YYYY-MM-DD 00:00:00"], ["RetireDate", "(optional) YYYY-MM-DD 23:59:59"],
        ["BackupSchema", "bck"], ["PartnerScoped", "N  (Y if the table has PartnerID)"],
        ["PartnerSource", "(if PartnerScoped=Y) SELECT TOP 1 PartnerID FROM dbo.tmfDefaults WITH (NOLOCK)"],
        ["NeverDelete", "N"],
    ]
    columns = [
        # Source: CELL = from the data tab | PARAM:x = script param | CONST:v = literal | ECHO:Col = copy | NULL
        {"ColumnName": "KeyCol", "SqlType": "varchar(20)", "Source": "CELL", "NullNormalize": "Y if blanks allowed"},
        {"ColumnName": "DataCol", "SqlType": "nvarchar(36)", "Source": "CELL", "NullNormalize": ""},
        {"ColumnName": "EffDate", "SqlType": "datetime", "Source": "PARAM:EffectiveDate", "NullNormalize": ""},
        {"ColumnName": "Flag", "SqlType": "char(1)", "Source": "CONST:Y", "NullNormalize": ""},
        {"ColumnName": "DecodeCol", "SqlType": "nvarchar(36)", "Source": "ECHO:DataCol", "NullNormalize": ""},
    ]
    operations = [{
        "Order": 1, "ActionTab": "MyData", "ActionFilter": "", "OpType": "INSERT (or UPDATE / DELETE)",
        "MatchKey": "KeyCol  (columns that uniquely identify a record)", "FromPredicate": "(UPDATE/DELETE only)",
        "SetMap": "(UPDATE only) Col <- CELL(New_Col)", "Idempotency": "NOT_EXISTS (or GUARDED / PATTERN)",
        "VerifyGroupBy": "(optional) column to group expected counts by",
    }]
    # The data tab holds ONE ROW PER RECORD with a column for each CELL field ONLY.
    # PARAM/CONST/ECHO/NULL columns are supplied by the tool -- do NOT put them in the tab.
    # (Add an optional 'Action' column + set ActionFilter only if the tab mixes
    #  actionable rows with reference rows you want skipped.)
    action = pd.DataFrame([["<key value>", "<data value>"]], columns=["KeyCol", "DataCol"])
    write_book(os.path.join(SAMPLES, "TEMPLATE.xlsx"), meta, columns, operations, {"MyData": action})


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--hts-source", help="work-item source workbook for the tmdHTSAdditional sample")
    args = ap.parse_args()
    os.makedirs(SAMPLES, exist_ok=True)
    build_template()
    build_globalcodes()
    if args.hts_source:
        build_htsadditional(args.hts_source)
    else:
        print("(skipped tmdHTSAdditional sample -- pass --hts-source to rebuild it)")
