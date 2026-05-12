# FA Rule 208 — Minimum Premiums (PolicyMinimumPremium_Ext)

## What this does

This creates an FA-only version of **Rule 208 Minimum Premiums** using the
`PolicyMinimumPremium_Ext` table from the FA ratebook.

**Source ratebook table** (cell B6 = `PolicyMinimumPremium_Ext`):

| MinimumPremiumType         | MinimumPremium |
|---------------------------|----------------|
| All Other                 | 1000           |
| Hired and Non-Owned Only  | 500            |

**Target rate page output:**
```
RULE 208. MINIMUM PREMIUMS
208.B. Rate and Premium Computation

    All other policies:                                              $1,000
    Policies providing hired auto and/or nonowned auto coverage only:  $500
```

This is **FA Situation 2** — same rule number as BA (208), different data table
and different display labels. We override the build method and the page rule
method in `FA/FARates.py`. No changes to `buildFAPages()` are needed because
`_page_rule_208` is already called there (inherited from BA).

---

## All code to add — FA/FARates.py

### Section A: Build Method

Add this inside the `Auto` class, in **Section A** (FA-only BUILD methods):

```python
def buildFAPolicyMinimumPremium(self, company):
    """
    FA Rule 208 — reads PolicyMinimumPremium_Ext and returns a 2-row DataFrame.

    Ratebook columns: MinimumPremiumType | MinimumPremium
    Output columns:   Description        | Minimum Premium

    Row mapping:
        "All Other"                → "All other policies:"
        "Hired and Non-Owned Only" → "Policies providing hired auto and/or
                                      nonowned auto coverage only:"
    """
    raw = self.rateTables[company].get("PolicyMinimumPremium_Ext")

    if raw is None:
        # Table not present for this company — return empty DataFrame
        return pd.DataFrame(columns=["Description", "Minimum Premium"])

    # raw is a list-of-lists; row 0 = column headers, rows 1+ = data
    df = pd.DataFrame(raw[1:], columns=raw[0])
    df = df.dropna(how="all").reset_index(drop=True)

    # Keep only the two columns we need
    df = df[["MinimumPremiumType", "MinimumPremium"]].copy()

    # Map the type code to the full display description
    type_display = {
        "All Other":                "All other policies:",
        "Hired and Non-Owned Only": "Policies providing hired auto and/or "
                                    "nonowned auto coverage only:",
    }
    df["MinimumPremiumType"] = (
        df["MinimumPremiumType"]
        .map(type_display)
        .fillna(df["MinimumPremiumType"])   # keep unmapped values as-is
    )

    # Convert MinimumPremium to numeric so the format method can apply $#,##0
    df["MinimumPremium"] = pd.to_numeric(df["MinimumPremium"], errors="coerce")

    # Rename columns for the rate page
    df.columns = ["Description", "Minimum Premium"]

    return df
```

---

### Section B: Page Rule Override

Add this inside the `Auto` class, in **Section B** (FA-only PAGE methods).
This overrides the BA `_page_rule_208` so FA uses our new build method:

```python
def _page_rule_208(self, RatePages):
    """
    FA override of Rule 208 — Minimum Premiums.

    Uses PolicyMinimumPremium_Ext (FA table) instead of BA's expense
    constant table. Layout is a single-table sheet with:
        Title:    RULE 208. MINIMUM PREMIUMS
        Subtitle: 208.B. Rate and Premium Computation
        2 data rows: description | dollar amount
    """
    self.compareCompanies("PolicyMinimumPremium_Ext")

    for CompanyTest in self.CompanyListDif:
        comp_name = self.extract_company_name(CompanyTest)

        # Set company suffix (blank when all companies share identical rates)
        self.title_company_name = CompanyTest
        if len(self.CompanyListDif) == 1:
            self.title_company_name = ""

        ws_name = "Rule 208 " + self.title_company_name
        title   = "RULE 208. MINIMUM PREMIUMS " + self.title_company_name

        # generateWorksheet:
        #   - Single table  → formatWorksheet() is called automatically
        #   - useHeader=False because the rate page shows no column header row,
        #     just the two data rows directly under the subtitle
        RatePages.generateWorksheet(
            ws_name,
            title,
            "208.B. Rate and Premium Computation",
            self.buildFAPolicyMinimumPremium(comp_name),
            False,    # useIndex
            False,    # useHeader — suppress column header row (not shown on rate page)
        )

        # Stamp the correct company names in the printed footer
        self.overideFooter(RatePages.getWB()[ws_name], CompanyTest)

        # Apply currency format and column widths for this rule
        self.formatRule208FA(RatePages.getWB()[ws_name])
```

---

### Format Method

Add this inside the `Auto` class right after `_page_rule_208` above:

```python
def formatRule208FA(self, ws):
    """
    Rule 208 FA: apply currency format and column widths.

    Why a custom format method?
    formatWorksheet() applies rateFormat (#,##0.000) to all data cells by
    default. Minimum premiums must display as "$1,000" not "1000.000", so
    we override the number format on column B after the standard format runs.

    Layout after generateWorksheet (subtitle is set, useHeader=False):
        Row 1: title   (bold)
        Row 2: subtitle (italic)
        Row 3: blank
        Row 4: "All other policies:"              | 1000
        Row 5: "Policies providing hired auto..." | 500
    """
    from config.constants import CURRENCY_FORMAT  # "$#,##0"

    # Wide description column, narrow dollar column
    ws.column_dimensions["A"].width = 58
    ws.column_dimensions["B"].width = 14

    # Re-apply currency format to ALL data rows in column B
    # (formatWorksheet applied rateFormat by default — this overwrites it)
    for row_idx in range(4, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=2)
        if cell.value is not None:
            cell.number_format = CURRENCY_FORMAT   # "$#,##0"
```

---

## Where to put everything — file map

```
FA/FARates.py
├── Section A (FA build methods)
│   └── def buildFAPolicyMinimumPremium(self, company)     ← ADD HERE
│
├── Section B (FA page methods)
│   └── def _page_rule_208(self, RatePages)                ← ADD HERE (overrides BA)
│
├── def formatRule208FA(self, ws)                          ← ADD HERE (after _page_rule_208)
│
└── buildFAPages()
    └── self._page_rule_208(RatePages)  ← ALREADY THERE — no change needed
```

> **Why no change to `buildFAPages()`?**
> `buildFAPages()` already calls `self._page_rule_208(RatePages)` because it was
> inherited from BA. Python will automatically pick up the FA override as soon as
> you add the method above.

---

## Page break — no change needed

The table has only 2 data rows and fits easily on one page.
`BApagebreaks.py` already has a `"Rule 208"` handler (or it falls back to
`fit_single_page`), which is correct. No addition to `FApagebreaks.py` needed.

---

## Complete code block (copy-paste ready)

Paste this entire block into `FA/FARates.py` inside the `Auto` class:

```python
# =========================================================================
# FA Rule 208 — Policy Minimum Premium
# Source table: PolicyMinimumPremium_Ext
# Situation 2: FA uses different data than BA for the same rule
# =========================================================================

def buildFAPolicyMinimumPremium(self, company):
    raw = self.rateTables[company].get("PolicyMinimumPremium_Ext")
    if raw is None:
        return pd.DataFrame(columns=["Description", "Minimum Premium"])

    df = pd.DataFrame(raw[1:], columns=raw[0])
    df = df.dropna(how="all").reset_index(drop=True)
    df = df[["MinimumPremiumType", "MinimumPremium"]].copy()

    type_display = {
        "All Other":                "All other policies:",
        "Hired and Non-Owned Only": "Policies providing hired auto and/or "
                                    "nonowned auto coverage only:",
    }
    df["MinimumPremiumType"] = (
        df["MinimumPremiumType"]
        .map(type_display)
        .fillna(df["MinimumPremiumType"])
    )
    df["MinimumPremium"] = pd.to_numeric(df["MinimumPremium"], errors="coerce")
    df.columns = ["Description", "Minimum Premium"]
    return df


def _page_rule_208(self, RatePages):
    self.compareCompanies("PolicyMinimumPremium_Ext")
    for CompanyTest in self.CompanyListDif:
        comp_name = self.extract_company_name(CompanyTest)
        self.title_company_name = CompanyTest
        if len(self.CompanyListDif) == 1:
            self.title_company_name = ""
        ws_name = "Rule 208 " + self.title_company_name
        RatePages.generateWorksheet(
            ws_name,
            "RULE 208. MINIMUM PREMIUMS " + self.title_company_name,
            "208.B. Rate and Premium Computation",
            self.buildFAPolicyMinimumPremium(comp_name),
            False,  # useIndex
            False,  # useHeader — no column header row on the rate page
        )
        self.overideFooter(RatePages.getWB()[ws_name], CompanyTest)
        self.formatRule208FA(RatePages.getWB()[ws_name])


def formatRule208FA(self, ws):
    from config.constants import CURRENCY_FORMAT
    ws.column_dimensions["A"].width = 58
    ws.column_dimensions["B"].width = 14
    for row_idx in range(4, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=2)
        if cell.value is not None:
            cell.number_format = CURRENCY_FORMAT
```

---

## Checklist before running

- [ ] `buildFAPolicyMinimumPremium` added to Section A of `FA/FARates.py`
- [ ] `_page_rule_208` override added to Section B (or below Section A) in `FA/FARates.py`
- [ ] `formatRule208FA` added right after `_page_rule_208` in `FA/FARates.py`
- [ ] `buildFAPages()` left unchanged (the call is already there)
- [ ] Table code `"PolicyMinimumPremium_Ext"` matches the exact value in cell B6 of the FA ratebook sheet
- [ ] Run the FA pipeline and verify the "Rule 208" tab shows:
  - Title: `RULE 208. MINIMUM PREMIUMS`
  - Subtitle: `208.B. Rate and Premium Computation`
  - Row 1: `All other policies:` → `$1,000`
  - Row 2: `Policies providing hired auto and/or nonowned auto coverage only:` → `$500`

---

## If the table code is wrong

If you get a `KeyError` or the sheet is blank, the table code in the ratebook
may have changed. To check:

1. Open the FA ratebook in Excel
2. Find the sheet with the minimum premium data
3. Look at cell **B6** — that exact string (including `_Ext`) is what you pass
   to `.get("...")` in the build method

Current code uses: `"PolicyMinimumPremium_Ext"`
