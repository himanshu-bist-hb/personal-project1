"""
generate_fa_package.py
======================
One-shot script that creates the FA (Farm Auto) package as a parallel
copy of the BA (Business Auto) package.

For each file in BA/, a renamed copy is written to FA/ with a global
text replacement: BA/Business Auto tokens are rewritten to FA/Farm Auto.

Mappings:
    BA/BARates.py        ->  FA/FARates.py
    BA/BARatePages.py    ->  FA/FARatePages.py
    BA/BApagebreaks.py   ->  FA/FApagebreaks.py
    BA/ExcelSettingsBA.py->  FA/ExcelSettingsFA.py

The script is idempotent — re-running overwrites the FA package.
After this, the user can hand-edit FA/* to specialize Farm Auto logic.
"""
from pathlib import Path
import re
import shutil


ROOT = Path(__file__).resolve().parent
BA_DIR = ROOT / "BA"
FA_DIR = ROOT / "FA"


# ---------------------------------------------------------------------------
# File rename map
# ---------------------------------------------------------------------------
RENAMES = {
    "BARates.py":         "FARates.py",
    "BARatePages.py":     "FARatePages.py",
    "BApagebreaks.py":    "FApagebreaks.py",
    "ExcelSettingsBA.py": "ExcelSettingsFA.py",
}


# ---------------------------------------------------------------------------
# Text replacements
# Order matters: longest / most-specific tokens first so they aren't shadowed
# by shorter ones (e.g. "BARates" before "BA").
# ---------------------------------------------------------------------------
REPLACEMENTS = [
    # Phrases — case-sensitive
    ("Business Auto",        "Farm Auto"),
    ("business auto",        "farm auto"),
    ("BUSINESS AUTO",        "FARM AUTO"),

    # File / asset names
    ("BA Input File",        "FA Input File"),
    ("BA Rate Pages",        "FA Rate Pages"),
    ("BA Small Market",      "FA Small Market"),
    ("BA Middle Market",     "FA Middle Market"),
    ("BA Exceptions",        "FA Exceptions"),
    ("BA Analytics",         "FA Analytics"),
    ("BA NAICS",             "FA NAICS"),
    ("BA CW Ratebook",       "FA CW Ratebook"),

    # Module names — must come BEFORE bare BA so they aren't half-replaced
    ("BApagebreaks",         "FApagebreaks"),
    ("BARatePages",          "FARatePages"),
    ("BARates",              "FARates"),
    ("ExcelSettingsBA",      "ExcelSettingsFA"),

    # Functions / identifiers carrying BA in their name
    ("buildBAPages",         "buildFAPages"),

    # Constant
    ("BA_INPUT_FILE",        "FA_INPUT_FILE"),
]

# After all string substitutions, apply word-boundary BA → FA for any
# stray tokens.  This catches things like "BA workbook" in comments.
WORD_BOUNDARY_BA = re.compile(r"\bBA\b")


def transform(text: str) -> str:
    for old, new in REPLACEMENTS:
        text = text.replace(old, new)
    text = WORD_BOUNDARY_BA.sub("FA", text)
    return text


def main():
    if not BA_DIR.is_dir():
        raise SystemExit(f"BA/ directory not found at {BA_DIR}")

    FA_DIR.mkdir(exist_ok=True)

    # Always rewrite __init__.py
    (FA_DIR / "__init__.py").write_text(
        '"""FA Rate Pages package: Farm Auto orchestration, rate engine, '
        'Excel factory, page-breaks."""\n',
        encoding="utf-8",
    )

    written = []
    for src_name, dst_name in RENAMES.items():
        src = BA_DIR / src_name
        dst = FA_DIR / dst_name
        if not src.is_file():
            print(f"  skip (missing): {src}")
            continue
        original = src.read_text(encoding="utf-8")
        rewritten = transform(original)
        dst.write_text(rewritten, encoding="utf-8")
        written.append((src_name, dst_name, len(rewritten)))

    print("FA package generated:")
    for s, d, n in written:
        print(f"  BA/{s:24s} ->  FA/{d:24s}  ({n:,} chars)")
    print(f"\nDone. {len(written)} files written to {FA_DIR}/")


if __name__ == "__main__":
    main()
