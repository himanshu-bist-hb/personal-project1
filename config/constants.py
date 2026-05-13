"""
config/constants.py
===================
Single source of truth for every project-wide constant.

RULE: If a value appears more than once across the project, it belongs here.

To change a file path, company name, margin, or format string:
  ► Change it HERE only.  Do not search for the raw value elsewhere.

HOW TO ADD A NEW LOB
  1. Add its Input File constant below (e.g.  GL_INPUT_FILE = "GL Input File.xlsx")
  2. Add its company codes + full names to the relevant dicts below (or create
     a new LOB-specific section if the company set is entirely different).
  3. Create GLRates.py and GLRatePages.py mirroring the BA equivalents.
  4. Wire into app.py.
"""

from pathlib import Path


# ===========================================================================
#  INPUT FILES
#  Change these to point at a different file — one place only.
# ===========================================================================

# The BA Input File is NOT optional.  It drives LCM mappings, state
# exceptions, base-rate sheet names, and more.  Must be in CWD or a
# full path must be given.
BA_INPUT_FILE: str = "BA Input File.xlsx"

# Farm Auto uses the SAME input file as BA.
# If FA ever gets its own input file, change this constant only.
FA_INPUT_FILE: str = BA_INPUT_FILE


# ===========================================================================
#  NETWORK PATHS  (shared across all LoB files)
# ===========================================================================

CW_RATEBOOK_DEFAULT    = Path(r"M:\Actshare\Com\BA\CW Ratebook\BA CW Ratebook.xlsx")
FA_CW_RATEBOOK_DEFAULT = Path(r"M:\Actshare\Com\FA\CW Ratebook\FA CW Ratebook.xlsx")
NAICS_FILE             = Path(r"M:\Actshare\Com\BA\CW Ratebook\BA NAICS Codes and Definitions.xlsx")
NAICS_SHEET         = "NAICSDescriptions"
NAICS_SKIP_ROWS     = list(range(11))        # rows 0-10 are headers/branding


# ===========================================================================
#  RATEBOOK SHEET STRUCTURE
# ===========================================================================

DETAIL_SHEET       = "Rate Book Details"
DETAIL_STATE_ROW   = 3     # zero-based .iloc row index for state name
DETAIL_STATE_COL   = 4     # zero-based .iloc col index for state name
DETAIL_DATE_ROW    = 7     # zero-based .iloc row index for effective date
DETAIL_DATE_COL    = 4     # zero-based .iloc col index for effective date
DATE_FMT           = "%m-%d-%Y"

DATA_START_ROW     = 12    # real rate data starts here in every sheet
SHEET_ID_CELL      = "B6"  # human-readable table name lives in this cell
SKIP_SHEET_SUFFIX  = "RR"  # sheets whose A1 ends with this are skipped


# ===========================================================================
#  COMPANY CODES  →  FULL LEGAL NAMES
#  Used by overideFooter() and anywhere a display name is needed.
#  Add new company here; nowhere else.
# ===========================================================================

COMPANY_NAMES: dict[str, str] = {
    "NAFF":  "Nationwide Affinity Insurance Company of America",
    "NACO":  "Nationwide Assurance Company",
    "NGIC":  "Nationwide General Insurance Company",
    "CCMIC": "Colonial County Mutual Insurance Company",
    "HICNJ": "Harleysville Insurance Company of New Jersey",
    "NICOF": "Nationwide Insurance Company of Florida",
    "NMIC":  "Nationwide Mutual Insurance Company",
    "AICOA": "Allied Insurance Company of America",
    "NICOA": "Nationwide Insurance Company of America",
    "NPCIC": "Nationwide Property Casualty Insurance Company",
    "NWAG":  "Nationwide Agribusiness Insurance Company",
}

# Short numeric codes used when tab names exceed Excel's 31-char limit.
COMPANY_NUMBERS: dict[str, str] = {
    "NAFF":  "21",
    "NACO":  "27",
    "NGIC":  "25",
    "CCMIC": "12",
    "HICNJ": "G",
    "NICOF": "40",
    "NMIC":  "01",
    "AICOA": "32",
    "NICOA": "07",
    "NPCIC": "28",
    "NWAG":  "AG",
}


# ===========================================================================
#  STATE ABBREVIATIONS
# ===========================================================================

STATE_ABBREVIATIONS: dict[str, str] = {
    "Alabama": "AL", "Alaska": "AK", "Arizona": "AZ", "Arkansas": "AR",
    "California": "CA", "Colorado": "CO", "Connecticut": "CT",
    "Delaware": "DE", "District of Columbia": "DC", "Florida": "FL",
    "Georgia": "GA", "Hawaii": "HI", "Idaho": "ID", "Illinois": "IL",
    "Indiana": "IN", "Iowa": "IA", "Kansas": "KS", "Kentucky": "KY",
    "Louisiana": "LA", "Maine": "ME", "Maryland": "MD",
    "Massachusetts": "MA", "Michigan": "MI", "Minnesota": "MN",
    "Mississippi": "MS", "Missouri": "MO", "Montana": "MT",
    "Nebraska": "NE", "Nevada": "NV", "New Hampshire": "NH",
    "New Jersey": "NJ", "New Mexico": "NM", "New York": "NY",
    "North Carolina": "NC", "North Dakota": "ND", "Ohio": "OH",
    "Oklahoma": "OK", "Oregon": "OR", "Pennsylvania": "PA",
    "Rhode Island": "RI", "South Carolina": "SC", "South Dakota": "SD",
    "Tennessee": "TN", "Texas": "TX", "Utah": "UT", "Vermont": "VT",
    "Virginia": "VA", "Washington": "WA", "West Virginia": "WV",
    "Wisconsin": "WI", "Wyoming": "WY",
}


# ===========================================================================
#  EXCEL FORMATTING DEFAULTS
#  All format methods in ExcelSettingsBA.py read from here.
#  Change a margin or font ONCE here — every sheet picks it up automatically.
# ===========================================================================

FONT_NAME         = "Arial"
FONT_SIZE         = 10            # pt
LEFT_MARGIN       = 0.25          # inches
RIGHT_MARGIN      = 0.25
TOP_MARGIN        = 1.25
BOTTOM_MARGIN     = 0.95
HEADER_MARGIN     = 0.5
FOOTER_MARGIN     = 0.25
RATE_FORMAT       = "#,##0.000"
CURRENCY_FORMAT   = "$#,##0"
THIN_BORDER_COLOR = "C1C1C1"
PRINT_TITLE_ROWS  = "1:3"

# The fixed left-header text that appears on every printed page.
# Change here if the manual section ever changes.
HEADER_LEFT_TEXT  = "Commercial Lines Manual: Division One - Automobile"
