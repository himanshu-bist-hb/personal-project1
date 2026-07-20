"""
audit_all_programs_split.py
============================
QA tool for the "All Programs" section of the BOP rating manual.

WHY THIS EXISTS
---------------
Every rating table AllProgramsPage.py prints under "All Programs" is there
because we intend the factor(s) to apply the same way to every BOP program
(Hab, Auto, Food, Retail, Office, Service, Wholesale) — Class_Code_Min in
the ratebook identifies which program a row belongs to (see the "Class
Codes" tab of BOP Input File.xlsx / bop_config.DEFAULT_CLASS_CODES).

In practice AllProgramsPage.py's build*() methods only ever SAMPLE one
representative Class_Code_Min per table (usually 20000=Auto, sometimes
40000=Food) instead of checking that every program's rows actually agree
— except for the three tables that already print separate Hab / non-Hab
tabs (Property Deductible, and both Windstorm/Hail Deductible tables).
CURRENT_CODE_ASSUMPTION below documents exactly what each table currently
assumes, straight from reading AllProgramsPage.py.

This script re-derives, straight from the ratebook data, how many distinct
factor patterns actually exist across Class_Code_Min for EVERY raw table
that feeds "All Programs" — so we can confirm or refute that assumption
per table, instead of trusting a single sampled class code. It reuses the
exact same sheet parsing BA.BARatePages.process_ratebook uses (B6 = table
name, data starts row 12) and the exact same NACO > NAFF > NICOF > NGIC >
CW nesting waterfall AllProgramsPage.AllPrograms.buildDataFrame uses, so
results line up with what the generated rate pages would actually show.

USAGE
-----
    python -m BOP.audit_all_programs_split --ngic "NGIC ratebook.xlsx" ^
        [--naco ...] [--naff ...] [--nicof ...] [--hicnj ...] [--cw ...] ^
        [--out "All Programs Split Audit.xlsx"]

Only --ngic is required (matches BOPRatePages.run()). Provide whichever
other company files this state's rate pages would actually nest, to match
production exactly — a table only found in a lower-level company file
(NACO/NAFF/NICOF) won't be checked unless that file is passed in.

OUTPUT
------
An Excel workbook with two sheets:
  - "Summary" — one row per raw table: what the code currently assumes,
    what the data actually shows (SINGLE / HAB_SPLIT / SPLIT_NOT_HAB /
    MULTI_SPLIT / NO_PROGRAM_DIMENSION / TABLE_NOT_FOUND), which programs
    cluster together, and a recommendation.
  - "Detail" — for every table that split 2-way-but-not-on-Hab or 3+ ways,
    which Class_Code_Min values ended up in which group.

Classification meanings:
  SINGLE              every program's rows are identical -> keep one section
  HAB_SPLIT           exactly two groups, one of them is Hab alone -> the
                       Hab / non-Hab split the code already uses elsewhere
                       is correct here too
  SPLIT_NOT_HAB       exactly two groups, but not split on the Hab boundary
                       -> needs a human look before splitting
  MULTI_SPLIT         three or more distinct patterns -> per the user's own
                       rule of thumb, this rating variable belongs in the
                       individual program section(s), not All Programs
  NO_PROGRAM_DIMENSION table has no (or only one) Class_Code_Min value
                       present, so a split can't be tested either way
  TABLE_NOT_FOUND     not present in any of the ratebooks supplied
"""

import argparse
import io
from typing import Dict, List, Optional, Tuple

import pandas as pd

from BA.BARatePages import load_ratebook, process_ratebook
from BOP.bop_config import load_bop_config


# Every buildDataFrame(...) call inside AllProgramsPage.py's sheetSpecs list
# — i.e. every raw table that feeds the "All Programs" page. Table codes are
# copied verbatim (spacing/underscores included) from AllProgramsPage.py.
ALL_PROGRAMS_TABLES: List[str] = [
    "BP7_Peril_Sprinkler_Discount",
    "BP7_Peril_Protection_Class",
    "BP7_Peril_Masonry_Veneer",
    "BP7_Peril_ValuationBasis",
    "BP7_Peril_AmountOfAnnualIncrease_Factor",
    "BP7_Peril_PropertyDeductible",
    "BP7_Peril_WH_Deductible_Factor",
    "BP7 Peril_WH_Deductible_Per_Building",
    "BP7_Peril_Burglar_Alarm_Factor",
    "BP7_Peril_Fire_Alarm_Factor",
    "BP7 Peril Building_Age_Modifier",
    "BP7_Peril_Building_Amt_Insurance",
    "BP7_Peril_BPP_Amt_Insurance",
    "BP7_Peril_Blanket_Insurance_Ind",
    "BP7_Peril_BCEG_Factor",
    "BP7_Peril_Tenants_Improvements_and_Betterments_Factor",
    "BP7_EBLimitsRelativityModifier",
    "BP7_EBDeductibleFactor",
    "BP7_Peril_Medical_Payments_Decreased_Limit",
    "BP7_Peril_TerritorialFactor",
]

# What AllProgramsPage.py currently ASSUMES for each table, read straight off
# the code (which Class_Code_Min it samples, if any) — printed next to our
# data-driven finding so a reviewer can see "code assumes X, data shows Y".
CURRENT_CODE_ASSUMPTION: Dict[str, str] = {
    "BP7_Peril_Sprinkler_Discount": "single (samples 20000 for Bldg/BPP, 40000 for Bus Inc)",
    "BP7_Peril_Protection_Class": "single, incl. Hab (samples 20000; FS Bus Inc samples 40000)",
    "BP7_Peril_Masonry_Veneer": "single, incl. Hab (no Class_Code_Min filter at all)",
    "BP7_Peril_ValuationBasis": "single, incl. Hab (no Class_Code_Min filter at all)",
    "BP7_Peril_AmountOfAnnualIncrease_Factor": "single, incl. Hab (no Class_Code_Min filter at all)",
    "BP7_Peril_PropertyDeductible": "ALREADY SPLIT: separate Hab (10000) / non-Hab (40000) tabs",
    "BP7_Peril_WH_Deductible_Factor": "ALREADY SPLIT: separate Hab (10000) / non-Hab (40000) tabs",
    "BP7 Peril_WH_Deductible_Per_Building": "ALREADY SPLIT: separate Hab (10000) / non-Hab (40000) tabs",
    "BP7_Peril_Burglar_Alarm_Factor": "single, incl. Hab (samples 40000)",
    "BP7_Peril_Fire_Alarm_Factor": "single, incl. Hab (no Class_Code_Min filter at all)",
    "BP7 Peril Building_Age_Modifier": "single, incl. Hab (Bldg/BPP sample 20000; Bus Inc samples 40000, FS only)",
    "BP7_Peril_Building_Amt_Insurance": "single, incl. Hab (samples 20000)",
    "BP7_Peril_BPP_Amt_Insurance": "single, incl. Hab (samples 20000)",
    "BP7_Peril_Blanket_Insurance_Ind": "single, incl. Hab (no Class_Code_Min filter at all)",
    "BP7_Peril_BCEG_Factor": "single, incl. Hab (no Class_Code_Min filter at all)",
    "BP7_Peril_Tenants_Improvements_and_Betterments_Factor": "single, incl. Hab (no Class_Code_Min filter at all)",
    "BP7_EBLimitsRelativityModifier": "single, incl. Hab (samples 40000)",
    "BP7_EBDeductibleFactor": "single, incl. Hab (samples 40000)",
    "BP7_Peril_Medical_Payments_Decreased_Limit": "single, incl. Hab (no Class_Code_Min filter at all)",
    "BP7_Peril_TerritorialFactor": "single, incl. Hab (no Class_Code_Min filter at all)",
}

HAB_CODE = 10000
# Column-name substrings (case-insensitive) that mark a column as a rated
# VALUE (e.g. PropertyDeductibleFactor, BldgTerritoryFactor,
# LimitRelativityModifier) rather than a rating KEY (e.g. DeductibleAmount,
# Peril TypeCode, ProtectionClass). Matches this codebase's naming convention.
VALUE_COL_MARKERS = ("factor", "modifier")
FLOAT_ROUND = 6  # tolerance for comparing factor values across programs


def _is_value_column(col: str) -> bool:
    lc = str(col).lower()
    return any(m in lc for m in VALUE_COL_MARKERS)


def build_nested_table(rate_tables: Dict[str, Optional[Dict]], table_code: str) -> Optional[pd.DataFrame]:
    """Same NACO > NAFF > NICOF > NGIC > CW waterfall as
    AllProgramsPage.AllPrograms.buildDataFrame, applied to the raw
    {table_code: [[header row], [data row], ...]} dicts process_ratebook
    returns per company."""
    for company in ("NACO", "NAFF", "NICOF"):
        tbl = rate_tables.get(company)
        if tbl and table_code in tbl:
            rows = tbl[table_code]
            return pd.DataFrame(rows[1:], columns=rows[0])
    ngic = rate_tables.get("NGIC")
    if ngic and table_code in ngic:
        rows = ngic[table_code]
        return pd.DataFrame(rows[1:], columns=rows[0])
    cw = rate_tables.get("CW")
    if cw and table_code in cw:
        rows = cw[table_code]
        return pd.DataFrame(rows[1:], columns=rows[0])
    return None


def classify_table(df: pd.DataFrame, class_codes: Dict[int, str]) -> Tuple[str, dict]:
    """
    Returns (classification, details).
    classification is one of: NO_PROGRAM_DIMENSION, SINGLE, HAB_SPLIT,
    SPLIT_NOT_HAB, MULTI_SPLIT.
    details holds the equivalence classes of Class_Code_Min found (each
    group's rows are byte-for-byte identical to each other, after rounding
    factor values) plus which programs were entirely missing from the table.
    """
    if "Class_Code_Min" not in df.columns:
        return "NO_PROGRAM_DIMENSION", {}

    value_cols = [c for c in df.columns if _is_value_column(c)]
    key_cols = [c for c in df.columns if c != "Class_Code_Min" and c not in value_cols]

    df = df.copy()
    df["Class_Code_Min"] = pd.to_numeric(df["Class_Code_Min"], errors="coerce")
    df = df[df["Class_Code_Min"].isin(class_codes.keys())]

    for c in value_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").round(FLOAT_ROUND)

    present_codes = sorted(int(c) for c in df["Class_Code_Min"].dropna().unique().tolist())
    if len(present_codes) <= 1:
        return "NO_PROGRAM_DIMENSION", {"present_codes": present_codes}

    # Canonical per-program subframe: keys + values only, sorted by key
    # columns so row order doesn't cause a false mismatch.
    subframes = {}
    for code in present_codes:
        sub = df[df["Class_Code_Min"] == code][key_cols + value_cols]
        sub = sub.sort_values(by=key_cols, kind="stable").reset_index(drop=True)
        subframes[code] = sub

    # Group class codes into equivalence classes by identical subframe.
    groups: List[List[int]] = []
    for code in present_codes:
        placed = False
        for g in groups:
            if subframes[code].equals(subframes[g[0]]):
                g.append(code)
                placed = True
                break
        if not placed:
            groups.append([code])

    missing_codes = sorted(set(class_codes.keys()) - set(present_codes))
    details = {
        "groups": groups,
        "group_programs": [[class_codes[c] for c in g] for g in groups],
        "present_codes": present_codes,
        "missing_codes": missing_codes,
        "key_cols": key_cols,
        "value_cols": value_cols,
    }

    if len(groups) == 1:
        classification = "SINGLE"
    elif len(groups) == 2:
        classification = "HAB_SPLIT" if [HAB_CODE] in groups else "SPLIT_NOT_HAB"
    else:
        classification = "MULTI_SPLIT"

    return classification, details


RECOMMENDATION = {
    "SINGLE": "keep as one All Programs section",
    "HAB_SPLIT": "split into Hab / non-Hab sections",
    "SPLIT_NOT_HAB": "2-way split found but NOT along the Hab boundary — review before splitting",
    "MULTI_SPLIT": "3+ distinct factor patterns — move to individual program section(s), not All Programs",
}


def load_and_audit(
    ngic,
    naco=None,
    naff=None,
    nicof=None,
    hicnj=None,
    cw=None,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Core entry point — accepts EITHER a file path string OR a file-like
    object (e.g. io.BytesIO from a Streamlit upload; load_ratebook already
    branches on hasattr(x, 'read')) for each ratebook, same calling
    convention BOPRatePages.run() uses. Only ngic is required.

    Returns (summary_df, detail_df) — no file I/O, so this is safe to call
    straight from a web UI as well as the CLI below.
    """
    cfg = load_bop_config()
    class_codes = cfg.class_codes

    companies = {"NGIC": ngic, "NACO": naco, "NAFF": naff, "NICOF": nicof, "HICNJ": hicnj, "CW": cw}
    rate_tables: Dict[str, Optional[Dict]] = {}
    for name, source in companies.items():
        if not source:
            rate_tables[name] = None
            continue
        ef = load_ratebook(source)
        _, tables = process_ratebook(name, ef)
        rate_tables[name] = tables

    summary_rows = []
    detail_rows = []

    for table_code in ALL_PROGRAMS_TABLES:
        df = build_nested_table(rate_tables, table_code)
        if df is None:
            summary_rows.append({
                "Table": table_code,
                "Current code assumption": CURRENT_CODE_ASSUMPTION.get(table_code, ""),
                "Finding": "TABLE_NOT_FOUND",
                "Programs": "",
                "Missing programs": "",
                "Notes": "not present in any provided ratebook",
            })
            continue

        classification, details = classify_table(df, class_codes)

        if classification == "NO_PROGRAM_DIMENSION":
            summary_rows.append({
                "Table": table_code,
                "Current code assumption": CURRENT_CODE_ASSUMPTION.get(table_code, ""),
                "Finding": "NO_PROGRAM_DIMENSION",
                "Programs": "",
                "Missing programs": "",
                "Notes": "table has no (or only 1) Class_Code_Min value present — cannot test a per-program split",
            })
            continue

        groups_desc = " | ".join(", ".join(gp) for gp in details["group_programs"])
        missing_desc = ", ".join(class_codes[c] for c in details["missing_codes"])

        summary_rows.append({
            "Table": table_code,
            "Current code assumption": CURRENT_CODE_ASSUMPTION.get(table_code, ""),
            "Finding": classification,
            "Programs": groups_desc,
            "Missing programs": missing_desc,
            "Notes": RECOMMENDATION[classification],
        })

        if classification in ("SPLIT_NOT_HAB", "MULTI_SPLIT"):
            for g, gp in zip(details["groups"], details["group_programs"]):
                detail_rows.append({
                    "Table": table_code,
                    "Program group": ", ".join(gp),
                    "Class_Code_Min values": ", ".join(str(c) for c in g),
                })

    return pd.DataFrame(summary_rows), pd.DataFrame(detail_rows)


def to_excel_bytes(summary_df: pd.DataFrame, detail_df: pd.DataFrame) -> bytes:
    """Render (summary_df, detail_df) as a two-sheet .xlsx in memory —
    used by both the CLI (writes to disk) and the Streamlit UI (offers a
    download button without touching the filesystem)."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        detail_df.to_excel(writer, sheet_name="Detail", index=False)
    return buf.getvalue()


def run_audit(
    ngic: str,
    naco: Optional[str] = None,
    naff: Optional[str] = None,
    nicof: Optional[str] = None,
    hicnj: Optional[str] = None,
    cw: Optional[str] = None,
    out_path: str = "All Programs Split Audit.xlsx",
) -> pd.DataFrame:
    """CLI wrapper — same as load_and_audit but also prints a text summary
    and saves the Excel report to out_path."""
    summary_df, detail_df = load_and_audit(ngic, naco, naff, nicof, hicnj, cw)

    with open(out_path, "wb") as f:
        f.write(to_excel_bytes(summary_df, detail_df))

    pd.set_option("display.max_columns", None)
    pd.options.display.width = None
    print(summary_df.to_string(index=False))
    print(f"\nSaved: {out_path}")
    return summary_df


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--ngic", required=True, help="NGIC ratebook path (required, same as BOPRatePages.run())")
    p.add_argument("--naco", help="NACO ratebook path (optional)")
    p.add_argument("--naff", help="NAFF ratebook path (optional)")
    p.add_argument("--nicof", help="NICOF ratebook path (optional)")
    p.add_argument("--hicnj", help="HICNJ ratebook path (optional)")
    p.add_argument("--cw", help="CW ratebook path (optional — falls back to nothing if omitted; the network default in BOPRatePages.run() is not used here)")
    p.add_argument("--out", default="All Programs Split Audit.xlsx", help="output .xlsx report path")
    args = p.parse_args()
    run_audit(args.ngic, args.naco, args.naff, args.nicof, args.hicnj, args.cw, args.out)


if __name__ == "__main__":
    main()
