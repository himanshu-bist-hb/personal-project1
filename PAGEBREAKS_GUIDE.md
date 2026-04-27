# Page-Break Rules Guide

How to define page-break / print-setup rules for any sheet in the BA Rate Pages workbook.

All page-break logic lives in **`BApagebreaks.py`**. You only ever edit that one file. Adding a rule is **two steps**:

1. Write a handler function.
2. Register it in the `SHEET_RULES` list.

That's it. Everything else — corruption fix, Index ordering, defaults — is already handled.

---

## Table of contents

1. [How the pipeline runs](#1-how-the-pipeline-runs)
2. [The 4 helper functions you'll use 95% of the time](#2-the-4-helper-functions-youll-use-95-of-the-time)
3. [Step-by-step: adding a new rule](#3-step-by-step-adding-a-new-rule)
4. [Recipes for common scenarios](#4-recipes-for-common-scenarios)
5. [The COM API cheatsheet (for anything the helpers don't cover)](#5-the-com-api-cheatsheet)
6. [Rule precedence and ordering](#6-rule-precedence-and-ordering)
7. [Defaults applied to every sheet](#7-defaults-applied-to-every-sheet)
8. [Common mistakes and gotchas](#8-common-mistakes-and-gotchas)
9. [End-to-end example: adding "Rule 999"](#9-end-to-end-example-adding-rule-999)

---

## 1. How the pipeline runs

```
buildBAPages (BARates.py + ExcelSettingsBA.py)
        |
        v   .xlsx written by openpyxl
        |
process_pagebreaks(xlsx_path, pdf_path)         (BApagebreaks.py)
        |
        +-- Phase 1: openpyxl
        |     - rename sheets > 31 chars
        |     - move "Index" to position 0
        |     - mark "Index" visible + active
        |     - openpyxl save
        |
        +-- Phase 2: Excel COM
              - Workbooks.Open(..., CorruptLoad=1)   <-- silent repair
              - for each sheet:
                    apply defaults  (fit-to-single-page, $1:$1 titles)
                    run matching SHEET_RULES handler
              - Sheets("Index").Activate()
              - SaveAs(temp, FileFormat=51)          <-- clean canonical xlsx
              - os.replace(temp, original)           <-- atomic swap
```

Result: a clean `.xlsx` that opens **without** the "Open and Repair" popup, with the Index tab at the front, and with rule-specific page breaks applied.

---

## 2. The 4 helper functions you'll use 95% of the time

These are defined at the top of `BApagebreaks.py`. Use them inside any handler.

| Helper | What it does | When to use |
|---|---|---|
| `fit_single_page(ws)` | Shrinks the whole sheet onto one printed page | Small/medium tables that should print on one page |
| `fit_width_only(ws)` | Width fits 1 page; height grows; manual breaks honored | Wide tables that span multiple pages tall |
| `disable_fit_to_page(ws)` | Turns off shrink-to-fit; manual breaks are honored as-is | Sheets where you control breaks manually |
| `add_break_after(ws, row)` | Inserts a horizontal page break **after** the given row | Anywhere you want a forced page boundary |

**Important:** if you use `add_break_after`, you almost always want `disable_fit_to_page` too — otherwise Excel may shrink the content small enough that the manual break has nothing to break across.

---

## 3. Step-by-step: adding a new rule

### Step 1 — write a handler function

Anywhere in the rule-handlers section of `BApagebreaks.py`. Signature is always:

```python
def _handle_rule_NAME(ws, app, dest_filename):
    ...
```

- `ws` — Excel COM Worksheet object (use `ws.PageSetup.*`, `ws.HPageBreaks`, etc.)
- `app` — Excel COM Application (used for `app.InchesToPoints(1.00)`)
- `dest_filename` — absolute path to the file (use this for state/file-specific branching)

### Step 2 — register the handler

Add one line to the `SHEET_RULES` list:

```python
SHEET_RULES = [
    ...
    ("Rule NAME", _handle_rule_NAME),
    ...
]
```

The first element is the **prefix** the sheet name has to start with. Order matters — see [section 6](#6-rule-precedence-and-ordering).

That's it. The next time `process_pagebreaks` runs, your rule fires for every sheet whose name starts with the prefix.

---

## 4. Recipes for common scenarios

### Scenario A — fit the entire sheet on one page

```python
def _handle_my_rule(ws, app, dest_filename):
    fit_single_page(ws)
```

### Scenario B — manual page break after a specific row

```python
def _handle_my_rule(ws, app, dest_filename):
    disable_fit_to_page(ws)
    add_break_after(ws, 37)
```

### Scenario C — break every N rows (running totals, repeating tables)

```python
def _handle_my_rule(ws, app, dest_filename):
    max_row = ws.UsedRange.Rows.Count
    disable_fit_to_page(ws)
    for row in range(45, max_row, 45):    # break every 45 rows
        add_break_after(ws, row)
```

### Scenario D — print area limited to specific columns

```python
def _handle_my_rule(ws, app, dest_filename):
    max_row = ws.UsedRange.Rows.Count
    ws.PageSetup.PrintArea = f"$A$1:$H${max_row}"
    fit_single_page(ws)
```

### Scenario E — repeat top rows on every printed page (title rows)

```python
def _handle_my_rule(ws, app, dest_filename):
    fit_width_only(ws)
    ws.PageSetup.PrintTitleRows = "$1:$3"   # repeat rows 1-3 at top of every page
```

### Scenario F — landscape orientation

```python
def _handle_my_rule(ws, app, dest_filename):
    ws.PageSetup.Orientation = XL_LANDSCAPE
    fit_width_only(ws)
```

### Scenario G — break BEFORE a row matching some text in column A

```python
def _handle_my_rule(ws, app, dest_filename):
    max_row = ws.UsedRange.Rows.Count
    disable_fit_to_page(ws)
    col_a = ws.Range(f"A1:A{max_row}").Value   # bulk read once = fast
    for row in range(1, max_row + 1):
        cell_value = str(col_a[row - 1][0])
        if cell_value.startswith("Section "):
            ws.HPageBreaks.Add(ws.Rows(row))   # break BEFORE this row
```

> **Why bulk-read?** Reading cells one at a time via COM is slow (one round-trip per cell). `ws.Range("A1:An").Value` returns a tuple of tuples in a single round-trip — orders of magnitude faster on large sheets.

### Scenario H — conditional: only apply if a specific cell has a value

```python
def _handle_my_rule(ws, app, dest_filename):
    if ws.Range("A10").Value != "Some Special Header":
        return                       # bail out; defaults stand
    fit_single_page(ws)
    ws.PageSetup.PrintTitleRows = "$1:$1"
```

### Scenario I — set a custom top margin

```python
def _handle_my_rule(ws, app, dest_filename):
    ws.PageSetup.TopMargin = app.InchesToPoints(1.00)   # 1.00 inch
    fit_single_page(ws)
```

### Scenario J — file-name-dependent behavior (e.g., FL has special stamping)

```python
def _handle_my_rule(ws, app, dest_filename):
    if "FL" not in dest_filename:
        ws.PageSetup.TopMargin = app.InchesToPoints(1.00)
    fit_single_page(ws)
```

---

## 5. The COM API cheatsheet

For anything the four helpers don't cover. All accessed via `ws.PageSetup.*` (or in a few cases `ws.*` directly).

### Print area / titles

```python
ws.PageSetup.PrintArea       = "$A$1:$H$50"     # rectangular range
ws.PageSetup.PrintTitleRows  = "$1:$3"          # repeat rows 1-3 on every page
ws.PageSetup.PrintTitleRows  = ""               # clear print titles
ws.PageSetup.PrintTitleColumns = "$A:$B"        # repeat columns A-B on every page
```

### Orientation

```python
ws.PageSetup.Orientation = XL_LANDSCAPE         # = 2
ws.PageSetup.Orientation = 1                    # portrait
```

### Margins (always pass through `app.InchesToPoints`)

```python
ws.PageSetup.TopMargin    = app.InchesToPoints(1.00)
ws.PageSetup.BottomMargin = app.InchesToPoints(0.95)
ws.PageSetup.LeftMargin   = app.InchesToPoints(0.25)
ws.PageSetup.RightMargin  = app.InchesToPoints(0.25)
ws.PageSetup.HeaderMargin = app.InchesToPoints(0.5)
ws.PageSetup.FooterMargin = app.InchesToPoints(0.25)
```

### Centering

```python
ws.PageSetup.CenterHorizontally = True
ws.PageSetup.CenterVertically   = True
```

### Fit-to-page (the helpers wrap these — use them directly only if you need something exotic)

```python
ws.PageSetup.Zoom             = False    # False = enable FitToPages mode
ws.PageSetup.Zoom             = 100      # any int = disable fit, use zoom %
ws.PageSetup.FitToPagesWide   = 1        # fit width to N pages
ws.PageSetup.FitToPagesTall   = 1        # fit height to N pages
ws.PageSetup.FitToPagesTall   = False    # height grows unlimited
```

### Page breaks

```python
ws.HPageBreaks.Add(ws.Rows(N))           # horizontal break ABOVE row N
                                         # (= page break AFTER row N-1)
ws.VPageBreaks.Add(ws.Columns(N))        # vertical break LEFT of column N
ws.ResetAllPageBreaks()                  # clear every manual break on the sheet
```

### Reading cells

```python
ws.Range("A10").Value                    # single cell read
ws.Range("A1:A50").Value                 # bulk read - returns tuple of tuples
ws.UsedRange.Rows.Count                  # how many rows have data
ws.UsedRange.Columns.Count               # how many columns have data
ws.Name                                  # sheet/tab name
```

### Writing cells (rare in handlers; usually buildBAPages did the writes)

```python
ws.Range("A1").Value = "Some title"
ws.Range("A1").Font.Bold = True
```

### Constants used in handlers

```python
XL_LANDSCAPE         = 2     # for PageSetup.Orientation
XL_OPEN_XML_WORKBOOK = 51    # for SaveAs FileFormat (xlsx)
XL_REPAIR_FILE       = 1     # for Workbooks.Open CorruptLoad
```

---

## 6. Rule precedence and ordering

Rules in `SHEET_RULES` are checked in order. The **first** matching prefix wins. So if you have two prefixes where one is a longer/more-specific version of the other, list the more specific one **first**:

```python
SHEET_RULES = [
    ("Rule 239 C",   _handle_rule_239c),        # specific  --- list FIRST
    ("Rule 239 ",    _handle_rule_239_general), # generic   --- list AFTER
]
```

If you swap the order, "Rule 239 C" would match `_handle_rule_239_general` and never reach `_handle_rule_239c`.

A sheet that doesn't match any prefix gets the [defaults](#7-defaults-applied-to-every-sheet) and nothing more.

---

## 7. Defaults applied to every sheet

Before any rule fires, every sheet receives these defaults:

```python
ws.PageSetup.PrintTitleRows = "$1:$1"   # row 1 repeats on each printed page
fit_single_page(ws)                     # whole sheet on one page
```

A handler can override either or both. Sheets without a registered rule keep these defaults — meaning **any unhandled sheet will print on a single page**, no matter how big the table is.

---

## 8. Common mistakes and gotchas

### Mistake 1 — manual breaks while fit-to-page is on

Excel honors manual breaks **only** when shrink-to-fit is off (or when fit allows multiple pages). If `fit_single_page` is in effect, your `add_break_after` calls will appear to do nothing because Excel is shrinking everything onto one page anyway.

> **Rule:** if you call `add_break_after`, also call `disable_fit_to_page` (or `fit_width_only` if you only want width to fit).

### Mistake 2 — listing a generic prefix before a specific one

`("Rule 222", ...)` will match `Rule 222 B`, `Rule 222 TTT`, and anything else starting with `Rule 222`. Always list specific prefixes first.

### Mistake 3 — `ws.HPageBreaks.Add(ws.Rows(N))` — off-by-one

`ws.Rows(N)` adds the break **above** row N, i.e. **after** row N-1. So:

| Want page to break after row… | Use |
|---|---|
| 37 | `ws.HPageBreaks.Add(ws.Rows(38))` *or* `add_break_after(ws, 37)` |
| Row before "Section X" | `ws.HPageBreaks.Add(ws.Rows(row_of_section_x))` |

`add_break_after(ws, row)` exists exactly so you don't have to think about this.

### Mistake 4 — reading hundreds of cells one at a time

```python
# SLOW: hundreds of COM round-trips
for row in range(1, max_row + 1):
    if str(ws.Range(f"A{row}").Value).startswith("X"):
        ...

# FAST: single round-trip
col_a = ws.Range(f"A1:A{max_row}").Value
for row in range(1, max_row + 1):
    if str(col_a[row - 1][0]).startswith("X"):
        ...
```

### Mistake 5 — touching openpyxl page-setup attributes in BApagebreaks

Don't. The whole point of routing page breaks through Excel COM is to avoid the corruption-popup issue that openpyxl-written page-setup XML triggers. The two phases of the pipeline are deliberately separated:

- **openpyxl phase** — sheet renames, sheet ordering, visibility. Nothing print-related.
- **COM phase** — every print/page-setup decision.

If you need a print-related setting you don't see in the helpers or the cheatsheet, look up the COM API for it — don't reach back to openpyxl.

### Mistake 6 — `ws.PageSetup.FitToPagesTall = 0`

Use `False` (which COM treats as "unlimited"), not `0`. Setting it to `0` may behave differently depending on Excel version. The helpers `fit_single_page` (=1) and `fit_width_only` (=False) get this right.

### Mistake 7 — not normalizing dest_filename for path checks

`dest_filename` may have backslashes, forward slashes, mixed case, etc. If you're checking for a state code like FL, do a substring check rather than path parsing:

```python
if "FL" in dest_filename:        # fine
if dest_filename.startswith(...): # fragile — careful with path normalization
```

---

## 9. End-to-end example: adding "Rule 999"

Suppose Rule 999 sheets need:
- print area `A1:K{max_row}`
- landscape orientation
- top margin 1 inch
- a manual page break every 50 rows
- title rows 1–2 repeated on every page
- but if cell `B4` says `"Special"`, just fit the whole thing on one page instead

**Step 1** — write the handler in `BApagebreaks.py`:

```python
def _handle_rule_999(ws, app, dest_filename):
    # Special-case branch
    if ws.Range("B4").Value == "Special":
        fit_single_page(ws)
        return

    max_row = ws.UsedRange.Rows.Count
    ws.PageSetup.PrintArea       = f"$A$1:$K${max_row}"
    ws.PageSetup.Orientation     = XL_LANDSCAPE
    ws.PageSetup.TopMargin       = app.InchesToPoints(1.00)
    ws.PageSetup.PrintTitleRows  = "$1:$2"
    disable_fit_to_page(ws)
    for row in range(50, max_row, 50):
        add_break_after(ws, row)
```

**Step 2** — register it in `SHEET_RULES`:

```python
SHEET_RULES = [
    ...
    ("Rule R1",    _handle_rule_r1),
    ("Rule 999",   _handle_rule_999),    # <-- new line
]
```

**Step 3** — restart Streamlit / re-run.

The next time `process_pagebreaks` runs, every sheet whose name starts with `"Rule 999"` will get this treatment. No other code needs to change.

---

## Where to find what

| Concern | File / location |
|---|---|
| All page-break rules | `BApagebreaks.py` (this is the only file you edit for page-break work) |
| The defaults (`fit_single_page`, `$1:$1` print titles) | Top of the for-loop in `process_pagebreaks` |
| Where the workbook is built | `BARates.py` (data) + `ExcelSettingsBA.py` (formatting) |
| Where `process_pagebreaks` is called | `BARatePages.py` (the `run()` function calls it after the openpyxl save) |
| State-wide constants (margins, fonts, default print titles) | `config/constants.py` |

---

## TL;DR

```python
# 1. Write a handler:
def _handle_rule_NEW(ws, app, dest_filename):
    fit_single_page(ws)            # or whatever you need
    add_break_after(ws, 37)        # if you want a manual break

# 2. Register it:
SHEET_RULES = [
    ...
    ("Rule NEW", _handle_rule_NEW),
]
```

Done.
