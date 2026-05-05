"""
SCRIPT 1 — Excel to JSON
Usage: python excel_to_json.py

Reads BA Input File.xlsx from the current directory and produces
ba_input_data.json — a single file you can copy and share anywhere.
Run json_to_excel.py on the other end to reconstruct the Excel file.
"""

import json
import os
import pandas as pd

INPUT_FILE  = "BA Input File.xlsx"
OUTPUT_FILE = "ba_input_data.json"


def excel_to_json(input_path, output_path):
    if not os.path.exists(input_path):
        print(f"ERROR: '{input_path}' not found in current directory.")
        return

    print(f"Reading: {input_path}")
    xl       = pd.ExcelFile(input_path, engine="openpyxl")
    result   = {}
    skipped  = []

    for sheet_name in xl.sheet_names:
        try:
            df = pd.read_excel(input_path, sheet_name=sheet_name,
                               header=0, engine="openpyxl")

            # Build list-of-lists: first row = column headers
            headers = [str(h) for h in df.columns.tolist()]
            rows    = [headers]

            for _, row in df.iterrows():
                clean_row = []
                for v in row.tolist():
                    if v is None:
                        clean_row.append(None)
                    elif isinstance(v, float) and pd.isna(v):
                        clean_row.append(None)
                    elif hasattr(v, "item"):          # numpy scalar
                        clean_row.append(v.item())
                    else:
                        clean_row.append(v)
                rows.append(clean_row)

            result[sheet_name] = rows
            print(f"  ✓ {sheet_name:35s}  ({len(rows)-1} data rows)")

        except Exception as e:
            skipped.append(sheet_name)
            print(f"  ✗ {sheet_name:35s}  SKIPPED — {e}")

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, default=str, ensure_ascii=False)

    size_kb = os.path.getsize(output_path) / 1024
    print(f"\nSaved: {output_path}  ({size_kb:.1f} KB)")
    print(f"Sheets exported : {len(result)}")
    if skipped:
        print(f"Sheets skipped  : {skipped}")
    print("\nYou can now copy ba_input_data.json and share it anywhere.")
    print("Use json_to_excel.py to rebuild the Excel file from it.")


if __name__ == "__main__":
    excel_to_json(INPUT_FILE, OUTPUT_FILE)
