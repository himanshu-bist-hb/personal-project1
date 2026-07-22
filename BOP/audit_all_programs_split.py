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
name, data starts row 12) and a lower-level-company > NGIC > CW nesting
waterfall — same idea as AllProgramsPage.AllPrograms.buildDataFrame, with
one deliberate difference: see LOWER_LEVEL_COMPANIES below for why HICNJ is
included here even though production's buildDataFrame currently skips it.

Every one of these tables carries a "Peril TypeCode" column (allother1,
cat1-4, fire1-4, ... plus a rollup value "allperil" — see the "bop checking
logic.jpeg" screenshot). Production's AllProgramsPage.py always filters
`Peril TypeCode` != "allperil" before sampling one Class_Code_Min and
assuming the result holds for every program — "allperil" rows are a
separate rollup consumed by the All Peril program page, and their factors
legitimately vary by Class_Code_Min for reasons that have nothing to do
with per-program consistency. So the check below is done PER Peril
TypeCode value, with "allperil" rows dropped before comparing: a single
peril that genuinely differs by program is flagged on its own, instead of
dragging every other (perfectly consistent) peril in the same table down
to a false split.

USAGE
-----
    python -m BOP.audit_all_programs_split --ngic "NGIC ratebook.xlsx" ^
        [--naco ...] [--naff ...] [--nicof ...] [--hicnj ...] [--cw ...] ^
        [--out "All Programs Split Audit.xlsx"]

Only --ngic is required (matches BOPRatePages.run()). Provide whichever
other company files this state's rate pages would actually nest, to match
production exactly — a table only found in a lower-level company file
(HICNJ/NICOF/NACO/NAFF) won't be checked unless that file is passed in.

OUTPUT
------
An Excel workbook with two sheets:
  - "Summary" — one row per (raw table, Peril TypeCode) found in that table
    (allperil rows excluded — see above): what the code currently assumes,
    what the data actually shows (SINGLE / HAB_SPLIT / SPLIT_NOT_HAB /
    MULTI_SPLIT / NO_PROGRAM_DIMENSION / ALLPERIL_ONLY / TABLE_NOT_FOUND),
    which programs cluster together, and a recommendation.
  - "Detail" — for every (table, peril) that split 2-way-but-not-on-Hab or
    3+ ways, which Class_Code_Min values ended up in which group.

Classification meanings:
  SINGLE              every program's rows are identical for this peril ->
                       keep in one All Programs section
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
  ALLPERIL_ONLY       every row for this table was the "allperil" rollup —
                       nothing left to compare once it's excluded; this
                       table belongs to the All Peril program page instead
  TABLE_NOT_FOUND     not present in any of the ratebooks supplied
"""

import argparse
import io
from typing import Dict, List, Optional, Tuple

import pandas as pd

from BA.BARatePages import load_ratebook, process_ratebook
from BOP.bop_config import load_bop_config
from config.constants import BOP_DETECT_ORDER

# Lower-level (company-specific) ratebooks that sit above NGIC (state
# default) / CW (country-wide fallback) in the nesting waterfall — same
# idea as Business Auto's LEVEL1 (company) -> LEVEL2 (NGIC) -> LEVEL3 (CW).
#
# IMPORTANT: these are PEERS, not a ranked hierarchy. In Business Auto each
# company builds its OWN separate output — LEVEL1 is always "this one
# company", checked against NGIC/CW; NACO is never checked against NAFF or
# NICOF because they're unrelated entities, not competitors. The same is
# true here: for any one state's BOP filing, exactly one of these companies
# is expected to be the one actually writing that business (a state files
# through HICNJ, or NICOF, or NACO, or NAFF — not several at once), so there
# is nothing to rank between them. build_nested_table() below treats them
# as a set, not an ordered list, and flags it rather than silently picking
# a winner if a table somehow turns up in more than one of them at once.
#
# The set of valid company codes still comes from BOP_DETECT_ORDER (the
# codebase's single list of BOP company codes) so this can't drift out of
# sync with it — only the ORDER from that list is deliberately not used
# here, since order would imply a ranking that doesn't exist.
LOWER_LEVEL_COMPANIES: List[str] = [c for c in BOP_DETECT_ORDER if c not in ("NGIC", "MM", "CW")]


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

# The peril-breakdown column every one of these tables carries (see the
# screenshot in "bop checking logic.jpeg"). "allperil" is a rollup row, not
# an individual peril -- AllProgramsPage.py itself always excludes it
# (`Peril TypeCode` != "allperil") before sampling one Class_Code_Min and
# assuming the result is the same for every program. Lumping "allperil" rows
# into the same comparison as the individual perils is what was causing
# every table to come back as a false MULTI_SPLIT: allperil's factors
# legitimately vary by Class_Code_Min for reasons unrelated to per-program
# consistency, so they must never be compared program-to-program here.
PERIL_COL = "Peril TypeCode"
ALLPERIL_VALUE = "allperil"

# Columns that identify WHICH Class_Code_Min band a row belongs to rather
# than an independent rating key -- Class_Code_Max is always Class_Code_Min
# + 9999 for these tables (e.g. 10000/19999 = Hab, 20000/29999 = Auto; see
# the screenshot), so it is mechanically different for every program by
# construction. Leaving it in the comparison "key" columns meant no two
# programs' rows could ever compare equal, which alone was enough to make
# every table come back split regardless of the allperil issue above.
CODE_RANGE_COLS = ("Class_Code_Min", "Class_Code_Max")


def _is_value_column(col: str) -> bool:
    lc = str(col).lower()
    return any(m in lc for m in VALUE_COL_MARKERS)


def build_nested_table(
    rate_tables: Dict[str, Optional[Dict]], table_code: str
) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """
    Peer lower-level company (whichever one was uploaded) -> NGIC -> CW
    waterfall, applied to the raw {table_code: [[header row], [data row],
    ...]} dicts process_ratebook returns per company. See the
    LOWER_LEVEL_COMPANIES comment above: HICNJ/NICOF/NACO/NAFF are peers,
    not a ranked chain, since only one of them is expected to be uploaded
    for any single state's run — same as Business Auto's per-company
    nesting (this company -> NGIC -> CW).

    Returns (df, note). note is None in the normal case; it's set to a
    short warning string if the table turned up in MORE THAN ONE
    lower-level company at once — an unexpected situation (it implies two
    different companies were both uploaded for the same state run) that
    this surfaces instead of silently resolving with an arbitrary pick.
    """
    hits = [c for c in LOWER_LEVEL_COMPANIES if rate_tables.get(c) and table_code in rate_tables[c]]

    note = None
    if len(hits) > 1:
        note = (
            f"found in multiple lower-level companies at once ({', '.join(hits)}) "
            f"— used {hits[0]}; normally only one lower-level company should be "
            f"uploaded per state, so this is worth double-checking"
        )

    if hits:
        tbl = rate_tables[hits[0]]
        rows = tbl[table_code]
        return pd.DataFrame(rows[1:], columns=rows[0]), note

    ngic = rate_tables.get("NGIC")
    if ngic and table_code in ngic:
        rows = ngic[table_code]
        return pd.DataFrame(rows[1:], columns=rows[0]), None
    cw = rate_tables.get("CW")
    if cw and table_code in cw:
        rows = cw[table_code]
        return pd.DataFrame(rows[1:], columns=rows[0]), None
    return None, None


def classify_group(df: pd.DataFrame, class_codes: Dict[int, str]) -> Tuple[str, dict]:
    """
    Classifies ONE peril group (all rows already narrowed to a single
    Peril TypeCode value, with that column dropped -- or the whole table,
    for the rare table that has no Peril TypeCode column at all) across
    Class_Code_Min (== program).

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
    key_cols = [c for c in df.columns if c not in CODE_RANGE_COLS and c not in value_cols]

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


def classify_table(df: pd.DataFrame, class_codes: Dict[int, str]) -> List[dict]:
    """
    Splits df by Peril TypeCode (dropping "allperil" rows entirely -- see
    the PERIL_COL comment above) and classifies each peril's rows
    independently across Class_Code_Min. This is the key fix: comparing one
    combined blob per table meant a single peril that genuinely differs by
    program (or the ever-present "allperil" rollup rows) dragged the WHOLE
    table down to a false MULTI_SPLIT, even when every other peril in that
    table was perfectly consistent across programs.

    Returns a list of {"peril": <value or None>, "classification": ...,
    "details": ...} dicts, one per Peril TypeCode value found (or a single
    entry with peril=None if the table has no Peril TypeCode column, or
    peril="(allperil only)" if every row was "allperil" and there is
    nothing else to check).
    """
    if "Class_Code_Min" not in df.columns:
        return [{"peril": None, "classification": "NO_PROGRAM_DIMENSION", "details": {}}]

    if PERIL_COL not in df.columns:
        classification, details = classify_group(df, class_codes)
        return [{"peril": None, "classification": classification, "details": details}]

    non_allperil = df[df[PERIL_COL].astype(str).str.strip().str.lower() != ALLPERIL_VALUE]
    peril_values = sorted(non_allperil[PERIL_COL].dropna().unique().tolist(), key=str)

    if not peril_values:
        return [{"peril": "(allperil only)", "classification": "ALLPERIL_ONLY", "details": {}}]

    results = []
    for peril in peril_values:
        sub = non_allperil[non_allperil[PERIL_COL] == peril].drop(columns=[PERIL_COL])
        classification, details = classify_group(sub, class_codes)
        results.append({"peril": peril, "classification": classification, "details": details})
    return results


RECOMMENDATION = {
    "SINGLE": "keep as one All Programs section",
    "HAB_SPLIT": "split into Hab / non-Hab sections",
    "SPLIT_NOT_HAB": "2-way split found but NOT along the Hab boundary — review before splitting",
    "MULTI_SPLIT": "3+ distinct factor patterns — move to individual program section(s), not All Programs",
    "ALLPERIL_ONLY": "every row is the \"allperil\" rollup — belongs to the All Peril program page, not the per-program All Programs check",
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
        df, nesting_note = build_nested_table(rate_tables, table_code)
        if df is None:
            summary_rows.append({
                "Table": table_code,
                "Peril TypeCode": "",
                "Current code assumption": CURRENT_CODE_ASSUMPTION.get(table_code, ""),
                "Finding": "TABLE_NOT_FOUND",
                "Programs": "",
                "Missing programs": "",
                "Notes": "not present in any provided ratebook",
            })
            continue

        for result in classify_table(df, class_codes):
            peril = result["peril"]
            classification = result["classification"]
            details = result["details"]
            peril_label = "" if peril is None else str(peril)

            if classification == "NO_PROGRAM_DIMENSION":
                note = "no (or only 1) Class_Code_Min value present — cannot test a per-program split"
                summary_rows.append({
                    "Table": table_code,
                    "Peril TypeCode": peril_label,
                    "Current code assumption": CURRENT_CODE_ASSUMPTION.get(table_code, ""),
                    "Finding": "NO_PROGRAM_DIMENSION",
                    "Programs": "",
                    "Missing programs": "",
                    "Notes": f"{note}; {nesting_note}" if nesting_note else note,
                })
                continue

            if classification == "ALLPERIL_ONLY":
                summary_rows.append({
                    "Table": table_code,
                    "Peril TypeCode": peril_label,
                    "Current code assumption": CURRENT_CODE_ASSUMPTION.get(table_code, ""),
                    "Finding": "ALLPERIL_ONLY",
                    "Programs": "",
                    "Missing programs": "",
                    "Notes": RECOMMENDATION["ALLPERIL_ONLY"],
                })
                continue

            groups_desc = " | ".join(", ".join(gp) for gp in details["group_programs"])
            missing_desc = ", ".join(class_codes[c] for c in details["missing_codes"])
            note = RECOMMENDATION[classification]

            summary_rows.append({
                "Table": table_code,
                "Peril TypeCode": peril_label,
                "Current code assumption": CURRENT_CODE_ASSUMPTION.get(table_code, ""),
                "Finding": classification,
                "Programs": groups_desc,
                "Missing programs": missing_desc,
                "Notes": f"{note}; {nesting_note}" if nesting_note else note,
            })

            if classification in ("SPLIT_NOT_HAB", "MULTI_SPLIT"):
                for g, gp in zip(details["groups"], details["group_programs"]):
                    detail_rows.append({
                        "Table": table_code,
                        "Peril TypeCode": peril_label,
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
