#!/usr/bin/env python3  
import argparse  
import pandas as pd  
from datetime import datetime  
from pathlib import Path  
from typing import Any, Optional
  
cols = [  
    "HTSNum","Chapter99","CountryofOrigin","StartEffDate","EndEffDate",  
    "TariffType","TariffGroup","RequiredStatusCode","ValidationLevel","ExportDate"  
]
  
def parse_args():  
    ap = argparse.ArgumentParser(description="Generate SQL from Excel action sheet")  
    ap.add_argument("--excel", required=True, help="Path to Excel file")  
    ap.add_argument("--sheet", required=True, help="Worksheet name to read")  
    ap.add_argument("--table", required=False, help="Target table name", default="[dbo].[tmdhtsAdditional]")  
    ap.add_argument("--out", required=True, help="Output SQL file")  
    return ap.parse_args()
  
def is_null(x: Any) -> bool:  
    if x is None:  
        return True  
    if isinstance(x, float) and pd.isna(x):  
        return True  
    if isinstance(x, str) and x.strip() == "":  
        return True  
    return bool(pd.isna(x)) if isinstance(x, (pd.Timestamp, pd.Series)) else False
  
def clean_number(n: Any) -> Optional[str]:  
    if is_null(n):  
        return None  
    # Many codes come in as floats like 850132.0 -> emit 850132 or keep decimals if needed  
    try:  
        fv = float(n)  
        iv = int(fv)  
        return f"N'{str(iv)}'" if fv == iv else f"N'{str(fv)}'"  
    except Exception:  
        return f"N'{str(n)}'"
  
def sql_literal(val: Any, col: Optional[str]) -> str:  
    if is_null(val):  
        return "N''"  
    # Datetime  
    if isinstance(val, pd.Timestamp):  
        s = val.strftime("%Y-%m-%d %H:%M:%S")  
        return f"CAST(N'{s}' AS DATETIME)"  
    # Numbers (strip .0 if integer)  
    if isinstance(val, (int, float)) and not isinstance(val, bool):  
        s = clean_number(val)  
        return "N''" if s is None else s  
    # Strings that might represent timestamps  
    if isinstance(val, str):  
        v = val.strip()  
        # Try parse datetime  
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y %H:%M", "%m/%d/%Y"):  
            try:  
                dt = datetime.strptime(v, fmt)  
                return f"CAST(N'{dt.strftime('%Y-%m-%d %H:%M:%S')}' AS DATETIME)"  
            except Exception:  
                pass  
        # Quote and escape  
        esc = v.replace("'", "''")  
        return f"N'{esc}'"  
    # Fallback
  
    if (col and col.find("Date")):  
        s = pd.Timestamp(val)  
        return f"CAST(N'{s}' AS DATETIME)"
  
    esc = str(val).replace("'", "''")  
    return f"N'{esc}'"
  
def col_or_null(row, col):  
    return None if col not in row or is_null(row[col]) else row[col]
  
def get_insert_command(table: str) -> str:  
    return f""" SELECT [RowId], [HTSNum], [Chapter99], [CountryofOrigin], [StartEffDate], [EndEffDate], [TariffType], [TariffGroup], [RequiredStatusCode], [ValidationLevel], [ExportDate] INTO {table} \n\t FROM (VALUES """
  
def make_insert(i, row, table: str) -> str:  
    present_cols = [c for c in cols if c in row]  
    values = []  
    for c in present_cols:  
        v = row[c]  
        values.append(sql_literal(v, c))  
    val_list = ", ".join(values)
  
    return f"\t\t,({val_list})"

def remove_first_comma(line: str) -> str:  
    """Remove the first comma from a line."""  
    comma_index = line.find(',')  
    if comma_index != -1:  
        return line[:comma_index] + line[comma_index + 1:]  
    return line  
  
def main():  
    args = parse_args()
  
    # Load Excel fast  
    df = pd.read_excel(args.excel, sheet_name=args.sheet, dtype=object)  
    # Normalize column names (strip spaces)  
    df.columns = [str(c).strip() for c in df.columns]
    df_clean = df[cols]
  
  
    df_clean = df_clean.drop_duplicates()   
    stmts = [get_insert_command(args.table)]  
    for _, row in df_clean.iterrows():  
        statement = make_insert(_, row, args.table)

        if _ == 0:
            statement = remove_first_comma(statement)

        if _ == len(df) - 1:
            statement += ');'           

        stmts.append(statement)        
  
    sql_text = "\n".join(stmts)  
    Path(args.out).write_text(sql_text, encoding="utf-8")  
    print(f"Wrote {len(stmts) - 1} statements to {args.out}")
  
if __name__ == "__main__":  
    main()  