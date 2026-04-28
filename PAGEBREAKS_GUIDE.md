# BA Rate Pages — Developer Guide

A complete, beginner-friendly walkthrough of how the project works and how to:

1. Add a new rule (a new tab in the output Excel).
2. Add a heading, subtitle, or multiple sub-tables to that tab.
3. Add a page-break rule so the printed pages look right.
4. Generate the PDF.

If you've never touched the codebase before, start at **Section 1**. Every step has a concrete copy-paste-able example.

---

## Table of contents

1. [The 30-second tour](#1-the-30-second-tour)
2. [The files you will touch](#2-the-files-you-will-touch)
3. [How a rate page is born — full lifecycle](#3-how-a-rate-page-is-born--full-lifecycle)
4. [PART 1 — Adding a new rule (new Excel tab)](#4-part-1--adding-a-new-rule-new-excel-tab)
5. [PART 2 — Headings, subtitles, sub-headings](#5-part-2--headings-subtitles-sub-headings)
6. [PART 3 — Page-break rules](#6-part-3--page-break-rules)
7. [PART 4 — PDF generation](#7-part-4--pdf-generation)
8. [Common mistakes (read this!)](#8-common-mistakes-read-this)
9. [Cheat sheet — where to find what](#9-cheat-sheet--where-to-find-what)

---

## 1. The 30-second tour

The app reads **Rate Books** (Excel files full of insurance rate tables) and produces a single **Rate Pages** Excel + PDF that gets filed with the Department of Insurance.

```
Streamlit UI (app.py)
        │  user uploads ratebooks, picks save folder, clicks Run
        ▼
BARatePages.run(...)            ← orchestration: opens files, calls everything
        │
        ▼
BARates.Auto.buildBAPages()     ← the BIG function — for each rule:
        │                          1. compares companies (clustering)
        │                          2. builds a DataFrame
        │                          3. tells the Excel factory to make a sheet
        ▼
ExcelSettingsBA.Excel           ← generateWorksheet*() creates the tab
        │                          formatWorksheet*() styles it
        ▼
   .xlsx is saved
        │
        ▼
BApagebreaks.process_pagebreaks ← applies page breaks / fit-to-page per rule
        │
        ▼
[user clicks "Generate PDF Document" in UI]
        ▼
BApagebreaks.export_to_pdf       ← drives Excel via COM to make a real PDF
        │
        ▼
   .pdf is saved next to the .xlsx
```

You will rarely touch `app.py` or `BARatePages.py`. The interesting work happens in **`BARates.py`**, **`ExcelSettingsBA.py`**, and **`BApagebreaks.py`**.

---

## 2. The files you will touch

| File | What it does | When to touch it |
|---|---|---|
| **`BARates.py`** | The "brain". Holds all rule logic, builds DataFrames, decides what goes on each sheet. Specifically, look inside `buildBAPages()`. | Adding a new rule, changing what data shows on a sheet, changing the title shown on the page. |
| **`ExcelSettingsBA.py`** | The "Excel factory". Knows how to draw a tab — title in A1, subtitle in A2, table starts at A4, borders, fonts, headers, footers. | Only if you need a **layout shape** that doesn't already exist (e.g. 17 stacked tables). 99% of the time you reuse the existing methods. |
| **`BApagebreaks.py`** | Post-processing. Decides where each printed page should break, what the print area is, what's portrait vs landscape, etc. | Whenever you add a new rule and want it to print the way your team expects. |
| **`config/constants.py`** | Single source of truth: margins, fonts, paths, company names. | Only when changing things that affect the **whole** workbook (e.g. font size 10 → 11). |

You will **not** touch:

- `BARatePages.py` — orchestration, already wired up.
- `app.py` — UI, already wired up.
- `BA Input File.xlsx` — read-only dependency.

---

## 3. How a rate page is born — full lifecycle

For one rule (say "Rule 208 Expense Constant"), here's exactly what happens:

```
1. buildBAPages() reaches the Rule 208 block
2. self.compareCompanies('ExpenseConstant_Ext')
        ─► figures out which companies share the same numbers
3. for each cluster of companies:
       a) self.buildExpenseConstant(comp_name)
              ─► returns a pandas DataFrame
       b) RatePages.generateWorksheet(
              wsTitle       = 'Rule 208',                      # Excel TAB name
              tableTitle    = 'RULE 208. EXPENSE CONSTANT',    # cell A1
              tableSubtitle = '208.B. Rate and Premium ...',   # cell A2
              df            = the DataFrame from step (a),
              useIndex      = False,
              useHeader     = True)
       c) self.overideFooter(ws, CompanyTest)
              ─► writes the right company names in the footer
4. After all rules done: RatePages.createIndex()
        ─► builds the Index tab with one hyperlink per rule
5. The workbook is saved as .xlsx
6. process_pagebreaks(xlsx_out, pdf_out)
        ─► runs through every sheet, applies page-break rules
7. (Optional) user clicks "Generate PDF Document"
        ─► export_to_pdf(xlsx, pdf) drives Excel and exports
```

The two things that actually need *your* attention when adding a rule are:

- **Step 3b** — picking the right `generateWorksheet*` method.
- **Step 6** — registering a page-break rule for the new tab.

That's it.

---

## 4. PART 1 — Adding a new rule (new Excel tab)

### 4.1. The pattern, from 30,000 feet

Every rule block in `buildBAPages()` (in `BARates.py`) looks like this:

```python
# Rule XXX
self.compareCompanies('SomeRateTableName_Ext')        # cluster companies
for CompanyTest in self.CompanyListDif:
    comp_name = self.extract_company_name(CompanyTest)
    self.title_company_name = CompanyTest
    if len(self.CompanyListDif) == 1:
        self.title_company_name = ""                  # only one cluster → no suffix

    RatePages.generateWorksheet(
        'Rule XXX ' + self.title_company_name,        # tab name
        'RULE XXX. SHORT DESCRIPTION ' + self.title_company_name,  # A1
        'XXX.A. Rate Computation',                    # A2 (subtitle)
        self.buildRuleXXX(comp_name),                 # the DataFrame
        False,    # useIndex
        True      # useHeader
    )
    self.overideFooter(RatePages.getWB()['Rule XXX ' + self.title_company_name], CompanyTest)
```

Six lines of boilerplate; only the bold parts change between rules. Once you understand this pattern, you've understood 95% of `buildBAPages`.

### 4.2. Step-by-step: adding "Rule 999 — Imaginary Coverage"

**Step 1 — Build the DataFrame.** Add a method to the `Auto` class (anywhere in `BARates.py`) that returns a `pd.DataFrame` shaped like the table you want to print:

```python
def buildRule999(self, comp_name):
    # Use self.buildDataFrame('SomeTableName_Ext') to fetch a nested rate table,
    # or build the DataFrame from scratch if it's a static table.
    df = self.buildDataFrame('ImaginaryFactor_Ext')
    df.columns = ['Coverage', 'Factor']
    return df
```

> **Tip:** look at `buildExpenseConstant`, `build231C`, `buildAntiqueAutoLiabFactors` — they're short, readable templates.

**Step 2 — Wire it into `buildBAPages()`.** Add a block following the existing pattern:

```python
# Rule 999
self.compareCompanies('ImaginaryFactor_Ext')
for CompanyTest in self.CompanyListDif:
    comp_name = self.extract_company_name(CompanyTest)
    self.title_company_name = CompanyTest
    if len(self.CompanyListDif) == 1:
        self.title_company_name = ""

    RatePages.generateWorksheet(
        'Rule 999 ' + self.title_company_name,
        'RULE 999. IMAGINARY COVERAGE ' + self.title_company_name,
        '999.A. Premium Computation',
        self.buildRule999(comp_name),
        False, True
    )
    self.overideFooter(RatePages.getWB()['Rule 999 ' + self.title_company_name], CompanyTest)
```

**Step 3 — Add a format pass at the bottom of `buildBAPages()`.** After all `generateWorksheet*` calls, there's a long block where each rule gets a styling pass. Add:

```python
Rule999_Sheets = [name for name in excel_Sheet_names if name.startswith('Rule 999')]
for Rule in Rule999_Sheets:
    self.formatWorksheet(AutoPages[Rule])    # or one of the format* variants below
```

**Step 4 — Add a page-break rule** (covered in [Part 3](#6-part-3--page-break-rules)).

That's the entire rule-creation flow.

### 4.3. Picking the right `generateWorksheet*` method

The Excel factory exposes one method per **layout shape**. Pick the one that matches what you're drawing:

| If your sheet has… | Use |
|---|---|
| 1 table | `generateWorksheet(title, A1, A2, df, useIndex, useHeader)` |
| 2 tables, **each with its own subtitle** | `generateWorksheet2tables(title, A1, A2, df1, A2_for_df2, df2, ...)` |
| 2–14 tables, **one shared subtitle**, stacked vertically | `generateWorksheet2tbls`, `generateWorksheet3tables`, `generateWorksheet4tables`, … `generateWorksheet14tables` |
| Many tables, **each with its own subtitle**, dynamic count | `generateWorksheetTablesX(title, A1, [subtitles], [dfs], ...)` |
| 2 side-by-side tables (Rule 23B style — left & right columns) | `generateWorksheet23B(...)` |
| Rule 222 fixed-row layout (very specific) | `generateRule222(...)` |
| Add a 2nd table to an already-created tab (Fleet style) | `ModifyFleetTable(sheetName, subtitle, df, ...)` |

> **Under the hood:** `generateWorksheet2tbls` through `generateWorksheet14tables` all call the same helper `generate_stacked_tables(...)`. They exist as named aliases so older code reads naturally. You can call `generate_stacked_tables` directly with a list of dataframes.

### 4.4. The two boolean parameters: `useIndex`, `useHeader`

These are passed straight through to pandas' `dataframe_to_rows`:

- `useIndex = True` → write the DataFrame's row index as the first column. Almost always **False** for rate pages (you don't want pandas' 0,1,2,3 column).
- `useHeader = True` → write the DataFrame's column names as a header row. Almost always **True**.

If unsure: `False, True`.

### 4.5. The `overideFooter` call (don't forget!)

Every `generateWorksheet*` already writes a default footer. But when companies are clustered (e.g. NGIC + NACO share a tab), the footer must list **the right names for that cluster**. `overideFooter` does that:

```python
self.overideFooter(RatePages.getWB()['Rule 999 ' + self.title_company_name], CompanyTest)
```

Skipping this call leaves the *default* footer (which assumes all companies). If you forget, the DoI filing will list the wrong names. **Always include it.**

---

## 5. PART 2 — Headings, subtitles, sub-headings

The Excel layout for *every* rate page tab is:

```
+----------------------------------------------------------+
| Row 1: tableTitle           ← A1, BOLD, big              |
| Row 2: tableSubtitle        ← A2, italic                 |
| Row 3: (blank)                                            |
| Row 4: column headers ┐                                   |
| Row 5+: data rows     ├── the DataFrame                   |
|                       ┘                                   |
+----------------------------------------------------------+
| Header (top of every printed page):                       |
|   Left:   "Commercial Lines Manual: Division One - Auto"  |
|   Center: "{State} - Rate Exceptions"                     |
|   Right:  "Effective: New {date}  Renewal {date}"         |
| Footer (bottom of every page): company names + page #     |
+----------------------------------------------------------+
```

You control the highlighted parts via the arguments to `generateWorksheet*`.

### 5.1. The main heading (A1)

That's the third argument to `generateWorksheet`:

```python
RatePages.generateWorksheet(
    'Rule 208',                              # tab name
    'RULE 208. EXPENSE CONSTANT',            # ← THIS goes in cell A1 (bold)
    '208.B. Rate and Premium Computation',   # subtitle
    df, False, True
)
```

Convention used across the project:

```
RULE NNN. SHORT DESCRIPTION
```

If you want the company suffix on the heading too (so the printed page shows "RULE 208. EXPENSE CONSTANT NGIC"), append `self.title_company_name`:

```python
'RULE 208. EXPENSE CONSTANT ' + self.title_company_name
```

### 5.2. The subtitle (A2)

That's the fourth argument:

```python
'208.B. Rate and Premium Computation'   # ← cell A2, italic
```

If you don't want a subtitle, pass exactly the string `' '` (a single space). The factory has special-case logic: when `tableSubtitle == ' '`, the data starts on row 3 instead of row 4. This keeps base-rate pages compact.

```python
RatePages.generateWorksheet('Rule 222 TTT BR', 'RULE 222. … BASE RATES', ' ', df, False, True)
#                                                                       ^^^
#                                                          no subtitle, table is closer to title
```

### 5.3. Multiple sub-headings (one per table)

Two scenarios:

**A) You know exactly how many tables — use `generateWorksheet2tables`:**

```python
RatePages.generateWorksheet2tables(
    'Rule 232 B',
    'RULE 232. PREMIUM DEVELOPMENT - PRIVATE PASSENGER TYPES',  # A1
    '232.B.1.b. Liability Fleet Size Factors',                  # subtitle for df1
    self.buildPPTLiabFleetFactors(comp_name),                   # df1
    '232.B.4.d. Physical Damage Fleet Size Factors',            # subtitle for df2
    self.buildPPTPhysDamFleetFactors(comp_name),                # df2
    False, True
)
```

The result:

```
RULE 232. PREMIUM DEVELOPMENT - PRIVATE PASSENGER TYPES
232.B.1.b. Liability Fleet Size Factors
   ┌────────────────────────────────┐
   │ ... df1 ...                    │
   └────────────────────────────────┘
232.B.4.d. Physical Damage Fleet Size Factors
   ┌────────────────────────────────┐
   │ ... df2 ...                    │
   └────────────────────────────────┘
```

**B) You have a *variable* number of sub-tables (or more than two) — use `generateWorksheetTablesX`:**

```python
subtitles = [
    '275.B.1.(a).(2). Trucks, Tractors, and Trailers Factors',
    '275.B.1.(a).(2). Private Passenger Types Factor',
    '275.B.1.(a).(2). Motorcycles Factors',
]
tables = [
    self.buildSomething1(comp_name),
    self.buildSomething2(comp_name),
    self.buildSomething3(comp_name),
]

RatePages.generateWorksheetTablesX(
    'Rule 275',
    'RULE 275. LEASING OR RENTAL CONCERNS',     # A1
    subtitles,                                  # list — one subtitle per df
    tables,                                     # list of dataframes
    False, True
)
```

This is the most flexible option. Use it whenever the number of sub-tables depends on state or company.

### 5.4. Plain stacked tables (no individual subtitles)

If you want N tables stacked under **one** shared subtitle:

```python
RatePages.generateWorksheet3tables(
    'Rule 222 B',
    'RULE 222. PREMIUM DEVELOPMENT - TRUCK, TRACTOR, TRAILER',
    '222.B.1.a. Liability Fleet Size Factors',          # one subtitle for the whole sheet
    df1, df2, df3,
    False, True
)
```

There's `generateWorksheet2tbls` … `generateWorksheet14tables` for 2–14 dataframes. They differ only in how many `df` arguments they accept.

### 5.5. Header & footer (top/bottom of every printed page)

These are applied automatically by the format methods. You don't need to set them explicitly. To change:

- **Center / right header text** (e.g. "Rate Exceptions" wording, effective dates) → edit `_apply_standard_header` in `ExcelSettingsBA.py`. They're set per workbook based on `self.State`, `self.nEffective`, `self.rEffective`.
- **Left header text** → edit `HEADER_LEFT_TEXT` in `config/constants.py` (changes for every sheet).
- **Footer company names** → handled by `overideFooter` in `BARates.py` per cluster.
- **Page number / state / tab name in center footer** → already wired (`{StateAbb} - &[Tab] - &P`). Editing it changes every sheet at once.

---

## 6. PART 3 — Page-break rules

This is `BApagebreaks.py`. You only ever edit this one file for page-break work.

### 6.1. The mental model

After the workbook is saved as `.xlsx`, `process_pagebreaks(xlsx_path, pdf_path)` opens it with **openpyxl** (no Excel app needed) and walks every sheet. For each sheet it:

1. Sets defaults: `print_title_rows = "1:1"` (row 1 repeats on every printed page) and `fit_single_page` (whole sheet on one page).
2. Walks the `SHEET_RULES` list and runs the **first** handler whose prefix matches the sheet name.

So a "page-break rule" is just a Python function that decides how a particular sheet should print. Two steps to add one:

1. Write the handler.
2. Register it in `SHEET_RULES`.

### 6.2. The 4 helpers you'll use 95% of the time

These are defined at the top of `BApagebreaks.py`. Use them inside any handler.

| Helper | What it does | When to use |
|---|---|---|
| `fit_single_page(ws)` | Shrinks the whole sheet onto one printed page | Small/medium tables that fit on one page |
| `fit_width_only(ws)` | Width fits 1 page; height grows; manual breaks honored | Tables that are wide but span multiple pages tall |
| `disable_fit_to_page(ws)` | Turns off shrink-to-fit entirely | When you want manual breaks to control everything |
| `add_break_after(ws, row)` | Inserts a horizontal page break **after** the given row | Wherever you want a forced page boundary |

> **Critical:** `add_break_after` only works if `fit_single_page` is OFF. If you call `add_break_after` you almost always also call `disable_fit_to_page` (or `fit_width_only`). Otherwise Excel shrinks everything onto one page and your break has no effect.

### 6.3. The handler signature

Every handler is a function with this exact shape:

```python
def _handle_rule_NAME(ws, dest_filename):
    ...
```

- `ws` — the openpyxl `Worksheet` object for the sheet being processed.
- `dest_filename` — absolute path of the `.xlsx` file. Useful for state-specific behavior (e.g. "do this only for FL files").

The function returns nothing. It mutates `ws` in place.

### 6.4. The openpyxl page-setup API (cheat sheet)

| What you want | Code |
|---|---|
| Set the print area | `ws.print_area = f"A1:H{ws.max_row}"` |
| Repeat top rows on every page | `ws.print_title_rows = "1:3"` |
| No print titles | `ws.print_title_rows = None` |
| Landscape orientation | `ws.page_setup.orientation = "landscape"` |
| Top margin (inches) | `ws.page_margins.top = 1.00` |
| Center horizontally | `ws.print_options.horizontalCentered = True` |
| Add page break after row N | `add_break_after(ws, N)`  *(or)*  `ws.row_breaks.append(Break(id=N+1))` |
| Read a cell | `ws["A10"].value` |
| Iterate rows | `for row in range(1, ws.max_row + 1):` |
| Read column A as a list | `[ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]` |

### 6.5. Step-by-step: adding "Rule 999"

Suppose Rule 999 sheets need:

- Print area `A1:H{max_row}`
- Manual page break every 50 rows
- Top margin 1 inch
- But if `B4` says `"Special"`, just fit on one page instead

**Step 1 — write the handler in `BApagebreaks.py`**, alongside the others:

```python
def _handle_rule_999(ws, dest_filename):
    # Special case: short sheet, just fit it
    if ws["B4"].value == "Special":
        fit_single_page(ws)
        return

    ws.print_area = f"A1:H{ws.max_row}"
    ws.page_margins.top = 1.00
    disable_fit_to_page(ws)
    for row in range(50, ws.max_row, 50):
        add_break_after(ws, row)
```

**Step 2 — register it in the `SHEET_RULES` list**:

```python
SHEET_RULES = [
    ...
    ("Rule R1",   _handle_rule_r1),
    ("Rule 999",  _handle_rule_999),    # ← new line
]
```

That's it. The next run, every sheet whose name starts with `"Rule 999"` gets this treatment.

### 6.6. Recipes for common scenarios

```python
# A) Whole sheet on one page
def _handle_my_rule(ws, dest_filename):
    fit_single_page(ws)

# B) Manual break after row 37
def _handle_my_rule(ws, dest_filename):
    disable_fit_to_page(ws)
    add_break_after(ws, 37)

# C) Break every 45 rows (running totals, repeating tables)
def _handle_my_rule(ws, dest_filename):
    disable_fit_to_page(ws)
    for row in range(45, ws.max_row, 45):
        add_break_after(ws, row)

# D) Restrict print area to columns A–H
def _handle_my_rule(ws, dest_filename):
    ws.print_area = f"A1:H{ws.max_row}"
    fit_single_page(ws)

# E) Repeat rows 1–3 at the top of every printed page
def _handle_my_rule(ws, dest_filename):
    fit_width_only(ws)
    ws.print_title_rows = "1:3"

# F) Landscape
def _handle_my_rule(ws, dest_filename):
    ws.page_setup.orientation = "landscape"
    fit_width_only(ws)

# G) Break BEFORE every row whose column-A text starts with "Section "
def _handle_my_rule(ws, dest_filename):
    disable_fit_to_page(ws)
    for row in range(1, ws.max_row + 1):
        value = str(ws.cell(row=row, column=1).value or "")
        if value.startswith("Section ") and row > 3:
            ws.row_breaks.append(Break(id=row - 1))   # break BEFORE row

# H) State-specific tweak (only for FL)
def _handle_my_rule(ws, dest_filename):
    fit_single_page(ws)
    if "FL" not in dest_filename:
        ws.page_margins.top = 1.00
```

### 6.7. Rule precedence — order matters

`SHEET_RULES` is checked in order. The **first** matching prefix wins. So always list more-specific prefixes **before** less-specific ones:

```python
SHEET_RULES = [
    ("Rule 239 C",   _handle_rule_239c),         # specific  ← FIRST
    ("Rule 239 ",    _handle_rule_239_general),  # generic   ← AFTER
]
```

If you swap the order, "Rule 239 C" sheets would match the generic handler and never hit the specific one.

A sheet that matches **no** rule keeps the defaults (single-page fit, row 1 repeats).

### 6.8. Bonus: the XML sanitize pass

After saving, `_sanitize_xlsx` opens the `.xlsx` zip and removes a few quirks that openpyxl leaves behind (invalid `definedName` entries with `$0` refs). Without this, Excel pops up "We found a problem with some content. Open and Repair?" on every file open. You don't need to do anything — it runs automatically. Just know it exists if you ever debug a "Open and Repair" popup.

---

## 7. PART 4 — PDF generation

The PDF flow lives in two places:

- `BApagebreaks.export_to_pdf(xlsx_path, pdf_path)` — the actual converter.
- `BARatePages.generate_pdf_only(xlsx_path, pdf_path, progress_callback)` — wrapper called by the UI.

### 7.1. How it works

1. The `.xlsx` already has all print settings baked in (page breaks, fit-to-page, print areas, headers, footers — all set during the page-breaks pass).
2. `export_to_pdf` launches Excel via COM (`win32com`), opens the workbook **read-only**, hides the `Index` tab so it doesn't appear in the PDF, then calls Excel's own `ExportAsFixedFormat` to produce the PDF.
3. The function then verifies the PDF actually exists on disk and is non-empty. If not, it raises `RuntimeError`.
4. The PDF is saved next to the `.xlsx`, with the same name, e.g.:

   ```
   ME 03-01-26 BA Small Market Rate Pages.xlsx
   ME 03-01-26 BA Small Market Rate Pages.pdf
   ```

### 7.2. How the UI uses it

In `app.py`:

- After Excel creation, the green "✓ Excel created" message appears with a **"Generate PDF Document"** button.
- Clicking the button calls `generate_pdf_only`. If it succeeds, a second green message appears: "✓ PDF created: {filename}".
- If it fails, the error is shown red and the button stays available so you can retry.

### 7.3. Requirements

- **Microsoft Excel must be installed** on the machine running the app. The COM bridge talks to a real Excel process.
- `pywin32` must be installed in the venv: `pip install pywin32`. Already done.
- Close the file in Excel before clicking "Generate PDF" — if the file is open elsewhere, Excel will refuse to overwrite.

### 7.4. Multi-state mode

In multi-state mode, ticking "Generate PDF for each state" runs the same `generate_pdf_only` once per state, saving each PDF into a `PDF/` subfolder inside your save location.

### 7.5. Customizing the PDF

The PDF format mirrors whatever the `.xlsx` print settings say. So:

- Want an extra page break? Add it in the page-breaks rule (Section 6).
- Want different margins? Set `ws.page_margins.*` in the page-breaks rule.
- Want landscape? `ws.page_setup.orientation = "landscape"` in the page-breaks rule.
- Want the Index page to **appear** in the PDF? Remove the `if sheet.Name == "Index": sheet.Visible = 0` block in `export_to_pdf`.

You **never** touch PDF-specific code to customize layout — every layout decision lives in the page-breaks layer. PDF export just respects what's already there.

---

## 8. Common mistakes (read this!)

### Mistake 1 — Manual breaks while fit-to-page is on

```python
# WRONG: fit_single_page shrinks everything to one page, the manual break does nothing
def _handle_my_rule(ws, dest_filename):
    fit_single_page(ws)
    add_break_after(ws, 37)    # ineffective!

# RIGHT: turn fit off first
def _handle_my_rule(ws, dest_filename):
    disable_fit_to_page(ws)
    add_break_after(ws, 37)
```

### Mistake 2 — Listing a generic rule prefix before a specific one

```python
# WRONG order: "Rule 222" eats "Rule 222 B" and "Rule 222 TTT"
SHEET_RULES = [
    ("Rule 222",     _handle_rule_222_generic),
    ("Rule 222 TTT", _handle_rule_222ttt),    # never runs!
]

# RIGHT: specific first
SHEET_RULES = [
    ("Rule 222 TTT", _handle_rule_222ttt),
    ("Rule 222 B",   _handle_rule_222b),
    ("Rule 222",     _handle_rule_222_generic),
]
```

### Mistake 3 — Forgetting `overideFooter`

If you don't call `overideFooter(ws, CompanyTest)` after `generateWorksheet*`, the tab keeps the *default* footer (which assumes all 4 companies). For a clustered tab, that's the wrong list of company names — the DoI filing will be wrong.

### Mistake 4 — Wrong `useIndex` value

```python
# WRONG: pandas writes 0,1,2,3,... as the first column
RatePages.generateWorksheet('Rule 999', 'TITLE', 'subtitle', df, True, True)

# RIGHT: drop the index
RatePages.generateWorksheet('Rule 999', 'TITLE', 'subtitle', df, False, True)
```

If unsure: **`useIndex=False, useHeader=True`**.

### Mistake 5 — `add_break_after` off-by-one

`add_break_after(ws, N)` adds a break **after** row N (so the next page starts on row N+1). It maps to `ws.row_breaks.append(Break(id=N+1))` under the hood (`Break(id=K)` puts the break **above** row K). Use the helper — don't manually construct `Break` objects unless you really need to.

### Mistake 6 — Using openpyxl page-setup attributes inside `BApagebreaks` while expecting a PDF

The PDF respects the `.xlsx`'s settings exactly, so this is rarely a problem. But: if you mutate the workbook **after** `process_pagebreaks` runs (e.g. in a Streamlit hook), those changes won't be picked up until you save again. The pipeline order is: build → save → page-break → save → PDF.

### Mistake 7 — Trying to generate a PDF without Excel installed

`export_to_pdf` requires Excel + `pywin32`. On a machine without Excel, it'll raise `pywintypes.com_error`. There is no fallback — Excel's PDF engine is the only thing that respects every print setting accurately.

### Mistake 8 — Editing `_apply_default_footer` directly when you only need a one-off change

That method runs for *every* sheet. If you only want to change one rule's footer, override it after `generateWorksheet*` is called:

```python
RatePages.generateWorksheet('Rule 999', 'TITLE', 'subtitle', df, False, True)
ws = RatePages.getWB()['Rule 999']
ws.oddFooter.left.text = "Custom footer text"   # one-off override
```

---

## 9. Cheat sheet — where to find what

| Concern | File / location |
|---|---|
| Add a new rule (new tab) | `BARates.py` → `buildBAPages()` |
| Change what data goes on a sheet | `BARates.py` → the `build*` method for that rule |
| Change tab title / subtitle | `BARates.py` → arguments to `generateWorksheet*` |
| Change layout shape (1 table → 3 tables) | `BARates.py` → switch to a different `generateWorksheet*` method |
| Add a new layout shape | `ExcelSettingsBA.py` → mimic `generate_stacked_tables` |
| Change cell styling (borders, fonts, column widths) | `ExcelSettingsBA.py` → the `format*` method matching your generator |
| Change top header text | `ExcelSettingsBA.py` → `_apply_standard_header` |
| Change left header text globally | `config/constants.py` → `HEADER_LEFT_TEXT` |
| Change default footer logic | `ExcelSettingsBA.py` → `_apply_default_footer` |
| Override footer for one cluster | `BARates.py` → call `self.overideFooter(ws, CompanyTest)` |
| Add page-break rule | `BApagebreaks.py` → write handler + add to `SHEET_RULES` |
| Change page-break defaults (apply to ALL sheets) | `BApagebreaks.py` → top of the for-loop in `process_pagebreaks` |
| PDF generation | `BApagebreaks.py` → `export_to_pdf`; called by `BARatePages.generate_pdf_only` |
| Fonts, margins, paths, company names | `config/constants.py` |
| State-specific sheet selection | `BA Input File.xlsx` (read by `sheet_fetch` in `buildBAPages`) |
| Streamlit UI behavior | `app.py` |

---

## TL;DR

**Adding a new rule:**

```python
# In BARates.py → buildBAPages()
self.compareCompanies('YourTable_Ext')
for CompanyTest in self.CompanyListDif:
    comp_name = self.extract_company_name(CompanyTest)
    self.title_company_name = CompanyTest
    if len(self.CompanyListDif) == 1:
        self.title_company_name = ""

    RatePages.generateWorksheet(                         # ← pick the right generator
        'Rule 999 ' + self.title_company_name,           # tab name
        'RULE 999. YOUR TITLE ' + self.title_company_name,  # cell A1
        '999.A. Your Subtitle',                          # cell A2
        self.buildRule999(comp_name),                    # the DataFrame
        False, True
    )
    self.overideFooter(RatePages.getWB()['Rule 999 ' + self.title_company_name], CompanyTest)

# At the bottom of buildBAPages, add a format pass:
Rule999_Sheets = [n for n in excel_Sheet_names if n.startswith('Rule 999')]
for r in Rule999_Sheets:
    self.formatWorksheet(AutoPages[r])
```

**Adding a page-break rule:**

```python
# In BApagebreaks.py
def _handle_rule_999(ws, dest_filename):
    disable_fit_to_page(ws)
    add_break_after(ws, 37)

SHEET_RULES = [
    ...
    ("Rule 999", _handle_rule_999),
]
```

**Generate PDF:** click the button. The new `export_to_pdf` does the rest — uses Excel's own engine, preserves every print setting, saves next to the xlsx, raises if anything fails.

Done.
