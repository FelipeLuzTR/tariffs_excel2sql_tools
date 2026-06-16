#!/usr/bin/env python3
"""Generic attachment -> standardized workbook adapter (config-driven; repeatable).

Turns a BA's raw delta attachment into the standardized _Meta / _Columns / _Operations
workbook that gen_dba_script.py consumes -- driven by a committed per-table PROFILE
(profiles/<table>.json) and a per-story SPEC (.json). Replaces hand-written, per-story
build scripts: the engineering interpretation lives in the small, diffable, reviewable
spec instead of throwaway python.

    python build_workbook.py --attachment <BA.xlsx> --spec <story.json> --out <workbook.xlsx>

PROFILE (profiles/<name>.json): `columns` (the _Columns sheet), `meta_defaults`,
`aliases` (extra source-column spellings), `default_keys` (match key per OpType).

SPEC (per story):
  {
    "profile": "tmdHTSAdditional",
    "meta": {"StoryId":"...", "Feature":"...", "Release":"26.3", "EffectiveDate":"...", "RetireDate":"..."},
    "operations": [
      {"optype":"DELETE", "tab":"DELETE"},                                   # keyed delete (key from profile)
      {"optype":"UPDATE", "tab":"UPDATE_EndEffDate",
       "where":{"column":"Chapter99","in":["99038212"]},                     # split one source tab into >1 op
       "set":"EndEffDate<-CELL(New_EndEffDate)",
       "guard":"t.EndEffDate > CAST(N'2026-06-07 23:59:59' AS DATETIME)"},
      {"optype":"INSERT", "tab":"INSERT", "verify_group":"Chapter99"},
      {"optype":"DELETE", "from_predicate":"<sql>"}                          # pattern delete (no tab)
    ]
  }
Per-op optional overrides: match_key, out_tab, idempotency.
"""
import argparse
import json
import os
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
PROFILES_DIR = os.path.join(HERE, "profiles")


def norm(c):
    return str(c).strip().lower().replace(" ", "")


def load_profile(name):
    with open(os.path.join(PROFILES_DIR, name + ".json"), encoding="utf-8") as f:
        return json.load(f)


def canon_map(profile):
    """norm(source column) -> canonical column name (from the profile columns + aliases)."""
    m = {norm(c["ColumnName"]): c["ColumnName"] for c in profile["columns"]}
    for canon, aliases in profile.get("aliases", {}).items():
        for a in aliases:
            m[norm(a)] = canon
    return m


def read_tab(attachment, tab, cmap):
    df = pd.read_excel(attachment, sheet_name=tab, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    ren = {c: cmap[norm(c)] for c in df.columns if norm(c) in cmap}
    return df.rename(columns=ren)


def apply_filter(df, where):
    if not where:
        return df
    col = where["column"]
    if col not in df.columns:
        raise SystemExit(f"filter column '{col}' not in tab columns {list(df.columns)}")
    s = df[col].map(lambda x: str(x).strip())
    if "in" in where:
        return df[s.isin({str(v).strip() for v in where["in"]})]
    if "equals" in where:
        return df[s == str(where["equals"]).strip()]
    raise SystemExit(f"filter for '{col}' needs 'in' or 'equals'")


def main():
    ap = argparse.ArgumentParser(description="Build a standardized DBA workbook from a BA attachment + spec.")
    ap.add_argument("--attachment", required=True, help="the BA's raw delta .xlsx")
    ap.add_argument("--spec", required=True, help="per-story build spec (.json)")
    ap.add_argument("--out", required=True, help="output standardized workbook (.xlsx)")
    args = ap.parse_args()

    with open(args.spec, encoding="utf-8") as f:
        spec = json.load(f)
    profile = load_profile(spec["profile"])
    cmap = canon_map(profile)
    cells = [c["ColumnName"] for c in profile["columns"] if c["Source"] == "CELL"]
    defkeys = profile.get("default_keys", {})

    # ---- _Meta (profile defaults <- spec.meta, plus the target table) ----
    md = dict(profile.get("meta_defaults", {}))
    md.update(spec.get("meta", {}))
    md["TargetTable"] = profile["target_table"]
    order = ["TargetTable", "StoryId", "Feature", "Release", "EffectiveDate", "RetireDate",
             "BackupSchema", "PartnerScoped", "PartnerSource", "NeverDelete"]
    meta_rows = [[k, md[k]] for k in order if k in md and str(md[k]).strip() != ""]

    # ---- _Operations + data tabs ----
    op_rows, tabs = [], {}
    default_idem = {"INSERT": "NOT_EXISTS", "UPDATE": "GUARDED"}
    for i, op in enumerate(spec["operations"], 1):
        optype = op["optype"].upper()
        has_tab = bool(op.get("tab"))
        out_tab = op.get("out_tab") or (f"{optype.title()}_{i}" if has_tab else "")
        match_key = op.get("match_key") or defkeys.get(optype, "")
        guard = op.get("from_predicate") or op.get("guard") or ""
        setmap = op.get("set", "")
        idem = op.get("idempotency") or default_idem.get(optype) or ("KEYED" if has_tab else "PATTERN")

        if has_tab:
            df = apply_filter(read_tab(args.attachment, op["tab"], cmap), op.get("where"))
            if optype == "INSERT":
                keep = list(cells)
            else:  # UPDATE / DELETE: the CELL key columns (+ the New_ source col for an UPDATE set)
                keep = [k.strip() for k in match_key.split(",") if k.strip() in cells]
                if setmap and "CELL(" in setmap:
                    keep.append(setmap.split("CELL(")[1].rstrip(")").strip())
            missing = [c for c in keep if c not in df.columns]
            if missing:
                raise SystemExit(f"op {i} ({optype}) tab '{op['tab']}' missing {missing}; have {list(df.columns)}")
            tabs[out_tab] = df[keep]
        elif optype == "DELETE" and not guard:
            raise SystemExit(f"op {i} DELETE has no 'tab' (keyed) and no 'from_predicate' (pattern).")

        op_rows.append({"Order": i, "ActionTab": out_tab, "ActionFilter": "", "OpType": optype,
                        "MatchKey": match_key, "FromPredicate": guard, "SetMap": setmap,
                        "Idempotency": idem, "VerifyGroupBy": op.get("verify_group", "")})

    # ---- write the standardized workbook ----
    with pd.ExcelWriter(args.out, engine="openpyxl") as w:
        pd.DataFrame(meta_rows).to_excel(w, sheet_name="_Meta", header=False, index=False)
        pd.DataFrame(profile["columns"])[["ColumnName", "SqlType", "Source", "NullNormalize"]].to_excel(
            w, sheet_name="_Columns", index=False)
        pd.DataFrame(op_rows)[["Order", "ActionTab", "ActionFilter", "OpType", "MatchKey",
                               "FromPredicate", "SetMap", "Idempotency", "VerifyGroupBy"]].to_excel(
            w, sheet_name="_Operations", index=False)
        for name, df in tabs.items():
            df.to_excel(w, sheet_name=name, index=False)

    print(f"wrote {args.out}")
    print(f"  profile={spec['profile']}  table={profile['target_table']}  ops={len(op_rows)}")
    for r in op_rows:
        n = len(tabs[r["ActionTab"]]) if r["ActionTab"] in tabs else 0
        print(f"   {r['Order']}. {r['OpType']:6} tab={r['ActionTab'] or '(pattern)'}  rows={n}  key={r['MatchKey']}")


if __name__ == "__main__":
    main()
