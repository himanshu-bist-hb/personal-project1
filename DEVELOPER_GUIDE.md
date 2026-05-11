# Developer Guide — BA / FA Rate Pages

This guide is written for people who maintain this tool but may not be software
engineers.  It explains how the system is organized and covers the most common
tasks: adding a rule, removing a rule, and changing a rule.

---

## What this tool does (plain English)

1. You provide one or more **ratebook Excel files** (the source of truth for
   rates) plus the **BA Input File.xlsx** (lookup tables, state exceptions).
2. The tool reads all the rates, applies the nesting and LCM rules, and writes
   a new **Rate Pages Excel file** formatted exactly as the Department of
   Insurance requires.
3. The Rate Pages file is then exported to **PDF** for filing.

---

## File map — one line per important file

| File | What it does |
|------|--------------|
| `app.py` | The Streamlit web UI — what users see and click |
| `BARatePages.py` | Top-level orchestrator: opens files, calls BA engine |
| `BA/BARates.py` | **Core BA engine** — all the rate-building logic lives here |
| `BA/ExcelSettingsBA.py` | Excel formatting factory — makes worksheets look right |
| `BA/BApagebreaks.py` | Adds page breaks / print settings to the saved .xlsx |
| `FA/FARates.py` | Farm Auto engine — inherits everything from BA |
| `FA/ExcelSettingsFA.py` | FA Excel factory — inherits from BA, add FA overrides here |
| `FA/FApagebreaks.py` | FA page breaks — inherits BA rules, add FA ones here |
| `FARatePages.py` | Top-level orchestrator for Farm Auto |
| `config/constants.py` | **Single source of truth** for file paths, company names, fonts |
| `BA Input File.xlsx` | Required input data: LCM mappings, state exceptions, etc. |

---

## How rules work

Every rate page (one Excel tab) corresponds to a **rule** (e.g., "Rule 208",
"Rule 297").  Each rule is now an isolated Python method inside the `Auto` class
in `BA/BARates.py`.

The method is named `_page_rule_XXX` (e.g., `_page_rule_208`).

What a rule method does:
1. Calls `self.compareCompanies(...)` — figures out which companies have
   identical rates so they can be grouped onto one tab.
2. Loops over the company groups and calls `RatePages.generateWorksheet(...)` —
   writes the data into an Excel worksheet.
3. Calls `self.overideFooter(...)` — stamps the correct company names in the
   page footer.
4. Calls a `self.formatXxx(...)` method — applies fonts, borders, column widths.

The master list of rules — in the order they appear in the output file — lives
inside `buildBAPages()` at the bottom of `BA/BARates.py`.  **Farm Auto** has its
own master list inside `buildFAPages()` in `FA/FARates.py`.

---

## Most common tasks

### Task 1 — Add a brand-new BA rule

Suppose a new Rule 999 is needed.

**Step A: Write the data builder (optional)**

If Rule 999 pulls from a new ratebook table, add a `buildRule999` method in
`BA/BARates.py` (near the other `build...` methods):

```python
def buildRule999(self, company):
    data = self.rateTables[company].get("MyNewTable_Ext")
    # ... transform data into a pandas DataFrame ...
    return df
```

**Step B: Write the page rule method**

Add the following method in `BA/BARates.py`, just above `buildBAPages`:

```python
def _page_rule_999(self, RatePages):
    """Rule 999: My New Coverage."""
    self.compareCompanies("MyNewTable_Ext")
    for CompanyTest in self.CompanyListDif:
        comp_name = self.extract_company_name(CompanyTest)
        self.title_company_name = CompanyTest
        if len(self.CompanyListDif) == 1:
            self.title_company_name = ""  # all companies identical, no suffix needed
        ws_name = 'Rule 999 ' + self.title_company_name
        RatePages.generateWorksheet(
            ws_name,
            'RULE 999. MY NEW COVERAGE ' + self.title_company_name,
            '999.B. Premium Computation',
            self.buildRule999(comp_name), False, True
        )
        self.overideFooter(RatePages.getWB()[ws_name], CompanyTest)
        # Optional: apply custom formatting
        # self.formatRule999(RatePages.getWB()[ws_name])
```

**Step C: Register it in `buildBAPages`**

Open `buildBAPages()` in `BA/BARates.py` and add one line where you want Rule
999 to appear in the output:

```python
self._page_rule_999(RatePages)
```

**Step D: Add a page-break rule (if needed)**

If Rule 999 needs specific page-break behavior, open `BA/BApagebreaks.py` and
follow the instructions at the top of that file.  Short version:

1. Write a handler function at the top of the file.
2. Add one line to `SHEET_RULES`.

---

### Task 2 — Add a brand-new FA rule (FA-only, not in BA)

**Step A & B:** Same as BA — write `buildFARule999` and `_page_rule_fa_999` in
`FA/FARates.py` (Section A and Section B at the top of that file).

**Step C:** Add the call inside `buildFAPages()` in `FA/FARates.py`:

```python
# ── FA-only rules ──────────────────────────────────────────────────────
self._page_rule_fa_999(RatePages)   # ← add this
```

**Step D (page breaks):** Open `FA/FApagebreaks.py` and follow the instructions
at the top of that file to add an FA-specific page-break rule.

---

### Task 3 — Remove a rule (for BA or FA)

Simply **comment out** the rule's line in the master list:

In `buildBAPages()` (for BA):
```python
# self._page_rule_999(RatePages)   ← commented out = rule skipped
```

In `buildFAPages()` (for FA):
```python
# self._page_rule_999(RatePages)   ← commented out = rule skipped for FA
```

The rule method itself stays in the file (in case you need it later).  The tab
simply won't appear in the output.

---

### Task 4 — Change a rule's data (different numbers)

If a rule pulls from the ratebook but the numbers changed because of a manual
revision, **nothing in this code needs to change** — the ratebook file itself is
the source of truth.  The tool will pick up the new numbers automatically.

If the *structure* of the data (column names, table codes) changed, edit the
`buildXxx` method in `BA/BARates.py`.

If **FA needs different data** than BA for the same rule, override the build
method in `FA/FARates.py`:

```python
class Auto(_BABase):
    def buildExpenseConstant(self, company):
        # FA uses a different table code
        data = self.rateTables[company].get("FAExpenseConstant_Ext")
        ...
        return df
```

---

### Task 5 — Change a rule's layout / formatting

Each rule has a `formatXxx` method in `BA/BARates.py` (e.g., `format208`).
Edit that method to change fonts, borders, column widths, or cell merges.

If **FA needs a different layout** for a rule, add a `generateWorksheetFACustom`
or `formatFACustom` method to `FA/ExcelSettingsFA.py` (see the docstring at the
top of that file for the exact pattern).

---

### Task 6 — Change a page-break behavior

Open `BA/BApagebreaks.py`.  The instructions at the top of that file explain
exactly how to add, change, or remove page-break rules.  The same approach
applies to `FA/FApagebreaks.py` for FA-only overrides.

---

### Task 7 — Change a file path or company name

Open `config/constants.py` — all paths, company codes, and formatting defaults
live there.  Change the value **once**, and every file that uses it picks up the
change automatically.

---

## VA, FL, and other state exceptions

Many rules have a `if self.StateAbb == "VA":` (or FL, MI, etc.) branch inside
the `_page_rule_XXX` method.  If a state gets new or different rules:

- **Different data**: edit the branch inside the relevant `_page_rule_XXX` method.
- **New tab unique to one state**: add it to `_page_rule_state_specific` in
  `BA/BARates.py` (all the CT / KS / MI / ND / NJ / NV / RI rules live there).
- **VA subtitle changes**: update `_VA_RULE_SUBTITLES` near the top of the
  `Auto` class in `BA/BARates.py`.

---

## Common mistakes and how to avoid them

| Mistake | How to avoid it |
|---------|----------------|
| Two worksheets with the same name crash Excel | Make sure the `ws_name` string is unique; company suffix handles this automatically if you follow the pattern |
| Rule appears in BA but not FA | Either add the call to `buildFAPages()` in `FA/FARates.py`, or explicitly leave it out if FA doesn't have that rule |
| Page breaks look wrong in the PDF | Check `BA/BApagebreaks.py` — every rule has a handler there (or inherits the default single-page fit) |
| New table code not found → KeyError | Make sure `_Ext` suffix is correct and the ratebook sheet name matches exactly |
| `shared` dict KeyError in Rule 451 | Rules 293, 297, and 451 must ALL be called in `buildBAPages` / `buildFAPages` for Rule 451 to work correctly — don't remove 293 or 297 if 451 is present |

---

## Inheritance summary

```
BA/BARates.py  (Auto class)
  ├── buildBAPages()          ← BA master list
  ├── _page_rule_208()
  ├── _page_rule_222_ttt_base_rates()
  ├── ... (79 _page_rule methods total)
  ├── build208(), build222B(), ... (data builders)
  ├── format208(), format222B(), ... (formatters)
  └── _sheet_fetch(), _VA_RULE_SUBTITLES, ...

FA/FARates.py  (Auto class, inherits from BA)
  ├── buildFAPages()          ← FA master list (can differ from BA)
  ├── [all BA _page_rule methods are inherited automatically]
  └── Add FA-only _page_rule_fa_xxx methods here

BA/ExcelSettingsBA.py  (Excel class)
  └── generateWorksheet, generateWorksheetTablesX, format methods, ...

FA/ExcelSettingsFA.py  (Excel class, inherits from BA)
  └── Add FA-only generate/format methods here

BA/BApagebreaks.py     — BA page-break SHEET_RULES registry
FA/FApagebreaks.py     — FA_SHEET_RULES = copy of BA + FA additions
```

---

## Quick reference: add a rule in 3 steps

```
1. Write   def _page_rule_999(self, RatePages): ...   in BA/BARates.py
2. Add     self._page_rule_999(RatePages)             in buildBAPages() / buildFAPages()
3. Add     ("Rule 999", _handle_rule_999)             in BApagebreaks.py SHEET_RULES  (if needed)
```

That is the entire workflow.
