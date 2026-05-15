"""
FA/FARates.py
=============
Farm Auto rate pages.

Farm Auto inherits EVERYTHING from Business Auto. The only thing this file
does right now is expose the `Auto` class under the FA package so that
FARatePages.py can import it as `FA.Auto`.

HOW FA INHERITANCE WORKS
-------------------------
`class Auto(_BABase)` means:
  - All data-build methods (buildExpenseConstant, buildBaseRates, ...)  ← inherited
  - All page-rule methods (_page_rule_208, _page_rule_222b, ...)        ← inherited
  - All format methods (format222B, formatBaseRates, ...)               ← inherited
  - buildBAPages() and buildFAPages()                                   ← inherited / defined here

THREE SITUATIONS WHEN EDITING FA RULES
----------------------------------------

Situation 1: Rule is IDENTICAL in FA → do nothing. It is already inherited.

Situation 2: Rule DATA changes in FA (different numbers):
  Override the build method here, e.g.:
    def buildExpenseConstant(self, company):
        # FA-specific data logic
        ...

Situation 3: Entirely NEW FA-only rule:
  Step A – Write buildFAXxx(company) in Section A below.
  Step B – Write _page_rule_fa_xxx(RatePages) in Section B below.
  Step C – Add self._page_rule_fa_xxx(RatePages) inside buildFAPages() in Section C.

HOW TO REMOVE A RULE FOR FA
-----------------------------
In buildFAPages() below, comment out the line for that rule:
    # self._page_rule_208(RatePages)   ← commented out = rule skipped for FA

HOW TO REORDER RULES
----------------------
In buildFAPages(), move the rule's line up or down.
"""

import warnings
import pandas as pd

from BA.BARates import Auto as _BABase
from config.constants import BA_INPUT_FILE

# Suppress noise we don't care about at import time
warnings.simplefilter("ignore", DeprecationWarning)
warnings.simplefilter("ignore", FutureWarning)


class Auto(_BABase):
    """
    Farm Auto rate page generator.
    Inherits all rules and formatting from Business Auto.

    FA HIERARCHY DIFFERENCES vs BA
    --------------------------------
    BA:  Company ratebook → NGIC (state-level, mandatory) → BA CW ratebook
    FA:  Company ratebook → NWAG (state-level, mandatory) → FA CW ratebook

    These two class attributes override the BA defaults so that every method
    that references the state-level company (compareCompanies, nesting) uses
    NWAG automatically — without duplicating any method code.
    """

    # ── FA hierarchy: NWAG is the state-level company (replaces BA's NGIC) ──
    _STATE_LEVEL_COMPANY = "NWAG"
    _COMPANIES_CHECK = [
        "NWAG", "NACO", "NAFF", "CCMIC", "HICNJ",
        "NICOF", "NMIC", "AICOA", "NICOA", "NPCIC",
    ]

    # =========================================================================
    # FA nesting override
    # =========================================================================

    def nesting(self):
        """
        FA nesting — identical logic to BA but with NWAG as the Level 2
        (state-level) ratebook instead of NGIC.

        Level 1: individual FA company ratebook (NACO, NAFF, etc.)
        Level 2: NWAG — mandatory FA state-level ratebook
        Level 3: FA CW ratebook (loaded under the "CW" key by FARatePages.run)

        NWAG MUST be last in ratebook_names so it is fully built (own tables +
        CW cascade) before any other company's Level 2 cascade uses it.
        """
        ratebook_names = [
            "NAFF", "NACO", "NICOF", "CCMIC", "HICNJ",
            "NICOA", "AICOA", "NPCIC", "NMIC",
            "NWAG",   # ← Level 2 for FA; must be last
        ]

        available_names = [
            name for name in ratebook_names
            if self.rateTables.get(name) not in (None, "Not found")
        ]

        for name in available_names:
            name, LEVEL1 = self.process_ratebook(name, self.rateTables)
            self.rateTables[name] = LEVEL1

        # FA equivalent of BA's MM suppression: if NMIC (MM) is provided,
        # suppress NWAG pages (same logic BA uses to suppress NGIC with NMIC).
        if (self.rateTables.get("NMIC") is not None) or (self.rateTables.get("CCMIC") is not None):
            self.rateTables["NWAG"] = None

    # =========================================================================
    # Section A: FA-only BUILD methods
    # Add a method here when FA needs data tables that BA does not have.
    # Pattern:
    #   def buildFAXxx(self, company):
    #       data = self.rateTables[company]["SomeNewFATable_Ext"]
    #       # ... build and return a pandas DataFrame
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

    def buildFALayupFactors(self, company):
        # FA uses FarmLayUpFactor_Ext (2-column table: Months Laid Up | Factor).
        # BA uses AutoLayUpFactor_Ext (many vehicle-type columns).
        raw = self.rateTables[company].get("FarmLayUpFactor_Ext")
        if raw is None:
            return pd.DataFrame(columns=["Months Laid Up", "Factor"])

        df = pd.DataFrame(raw[1:], columns=raw[0])
        df = df.dropna(how="all").reset_index(drop=True)

        # Ratebook stores "7-12"; rate page shows "7+"
        df["Months Laid Up"] = df["Months Laid Up"].replace("7-12", "7+")

        # Sort into display order: 0-2, 3, 4, 5, 6, 7+
        _order = {"0-2": 0, "3": 1, "4": 2, "5": 3, "6": 4, "7+": 5}
        df["_sort"] = df["Months Laid Up"].map(_order)
        df = df.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)

        df["Factor"] = df["Factor"].astype(float).map(lambda x: f"{x:.3f}")
        return df

    def buildFAPollutionBaseRate(self, company):
        # FA Rule 303 — FarmPollutionLiabilityBaseRate_Ext
        # Ratebook: Constant | BaseRate  →  Rate page: Base Rate: | Per $25,000 Limit of Liability
        raw = self.rateTables[company].get("FarmPollutionLiabilityBaseRate_Ext")
        if raw is None:
            return pd.DataFrame(columns=["", "Per $25,000 Limit of Liability"])
        df = pd.DataFrame(raw[1:], columns=raw[0])
        df = df.dropna(how="all").reset_index(drop=True)
        base_rate = pd.to_numeric(df["BaseRate"].iloc[0], errors="coerce")
        return pd.DataFrame({
            "": ["Base Rate:"],
            "Per $25,000 Limit of Liability": [f"${base_rate:,.2f}"],
        })

    def buildFAPollutionFleetFactor(self, company):
        # FA Rule 303 — FarmPollutionLiabilityFleetFactor_Ext
        # Ratebook: PowerUnitMinimum | Factor  →  Rate page: Fleet Size (ranges) | Fleet Factor
        raw = self.rateTables[company].get("FarmPollutionLiabilityFleetFactor_Ext")
        if raw is None:
            return pd.DataFrame(columns=["Fleet Size", "Fleet Factor"])
        df = pd.DataFrame(raw[1:], columns=raw[0])
        df = df.dropna(how="all").reset_index(drop=True)
        _range_map = {0: "0-4", 5: "5-9", 10: "10-19", 20: "20-29", 30: "30+"}
        df["PowerUnitMinimum"] = (
            pd.to_numeric(df["PowerUnitMinimum"], errors="coerce")
            .apply(lambda v: _range_map.get(int(v), str(int(v))) if pd.notna(v) else v)
        )
        df["Factor"] = pd.to_numeric(df["Factor"], errors="coerce").map(lambda x: f"{x:.2f}")
        df = df.rename(columns={"PowerUnitMinimum": "Fleet Size", "Factor": "Fleet Factor"})
        return df

    # =========================================================================
    # Section B: FA-only PAGE methods
    # Add a method here when FA needs a new rule page that BA does not have.
    # Pattern:
    #   def _page_rule_fa_xxx(self, RatePages):
    #       self.compareCompanies("SomeNewFATable_Ext")
    #       for CompanyTest in self.CompanyListDif:
    #           comp_name = self.extract_company_name(CompanyTest)
    #           self.title_company_name = CompanyTest
    #           if len(self.CompanyListDif) == 1:
    #               self.title_company_name = ""
    #           RatePages.generateWorksheet(
    #               "Rule FA XXX " + self.title_company_name,
    #               "RULE FA XXX. DESCRIPTION " + self.title_company_name,
    #               "FA XXX.B. Subtitle",
    #               self.buildFAXxx(comp_name), False, True,
    #           )
    #           self.overideFooter(
    #               RatePages.getWB()["Rule FA XXX " + self.title_company_name],
    #               CompanyTest,
    #           )
    # =========================================================================

    def _page_rule_208(self, RatePages):
        """FA override — Rule 208 uses PolicyMinimumPremium_Ext instead of BA's table."""
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
        from openpyxl.styles import Font, Alignment, Border, Side

        ws.column_dimensions["A"].width = 58
        ws.column_dimensions["B"].width = 14

        border = Border(
            left=Side(border_style="thin", color="C1C1C1"),
            right=Side(border_style="thin", color="C1C1C1"),
            top=Side(border_style="thin", color="C1C1C1"),
            bottom=Side(border_style="thin", color="C1C1C1"),
        )

        for row_idx in range(4, ws.max_row + 1):
            cell_a = ws.cell(row=row_idx, column=1)
            if cell_a.value is not None:
                cell_a.font      = Font(name="Arial", size=10)
                cell_a.alignment = Alignment(horizontal="left", vertical="center")
                cell_a.border    = border
            cell_b = ws.cell(row=row_idx, column=2)
            if cell_b.value is not None:
                cell_b.font          = Font(name="Arial", size=10)
                cell_b.alignment     = Alignment(horizontal="center", vertical="center")
                cell_b.border        = border
                cell_b.number_format = CURRENCY_FORMAT

    def _page_rule_222e(self, RatePages):
        # FA override — uses FarmLayUpFactor_Ext (2-col) instead of AutoLayUpFactor_Ext (multi-col).
        # Subtitle changes from "Commercial Lay-up Credit" to "Farm Lay-up Credit".
        self.compareCompanies("FarmLayUpFactor_Ext")
        for CompanyTest in self.CompanyListDif:
            comp_name = self.extract_company_name(CompanyTest)
            self.title_company_name = CompanyTest
            if len(self.CompanyListDif) == 1:
                self.title_company_name = ""
            tables    = [self.buildFALayupFactors(comp_name)]
            subtitles = ["222.E. Farm Lay-up Credit"]
            RatePages.generateWorksheetTablesX(
                "Rule 222 E" + self.title_company_name,
                "RULE 222.E PREMIUM DEVELOPMENT - TRUCK, TRACTOR, TRAILER TYPES " + self.title_company_name,
                subtitles, tables, False, True,
            )
            self.overideFooter(RatePages.getWB()["Rule 222 E" + self.title_company_name], CompanyTest)
        _wb = RatePages.getWB()
        for _sn in _wb.sheetnames:
            if _sn.startswith("Rule 222 E"):
                self.format222E(_wb[_sn])

    def _page_rule_225d(self, RatePages):
        # FA override — layup table changes; fleet-size table is the same.
        self.compareCompanies("FarmLayUpFactor_Ext")
        for CompanyTest in self.CompanyListDif:
            comp_name = self.extract_company_name(CompanyTest)
            self.title_company_name = CompanyTest
            if len(self.CompanyListDif) == 1:
                self.title_company_name = ""
            subtitles = [
                "225.D Fleet Size Rating Factors - Zone Rated",
                "225.D. Farm Lay-up Credit",
            ]
            tables = [
                self.buildFleetSizeRatingFactorsZone(comp_name),
                self.buildFALayupFactors(comp_name),
            ]
            title_start = "RULE 225.​​​ PREMIUM DEVELOPMENT - ZONE-RATED AUTOS "
            RatePages.generateWorksheetTablesX(
                "Rule 225.D " + self.title_company_name,
                title_start + self.title_company_name,
                subtitles, tables, False, True,
            )
            self.overideFooter(RatePages.getWB()["Rule 225.D " + self.title_company_name], CompanyTest)
        _wb = RatePages.getWB()
        for _sn in _wb.sheetnames:
            if _sn.startswith("Rule 225.D"):
                self.format225D(_wb[_sn])

    def _page_rule_239d(self, RatePages):
        # FA override — uses FarmLayUpFactor_Ext and Farm subtitle.
        self.compareCompanies("FarmLayUpFactor_Ext")
        for CompanyTest in self.CompanyListDif:
            comp_name = self.extract_company_name(CompanyTest)
            self.title_company_name = CompanyTest
            if len(self.CompanyListDif) == 1:
                self.title_company_name = ""
            title_start = "RULE 239.​ PUBLIC AUTO PREMIUM DEVELOPMENT - OTHER THAN ZONE-RATED AUTOS "
            RatePages.generateWorksheet(
                "Rule 239 D " + self.title_company_name,
                title_start + self.title_company_name,
                "239.D. Farm Lay-up Credit",
                self.buildFALayupFactors(comp_name), False, True,
            )
            self.overideFooter(RatePages.getWB()["Rule 239 D " + self.title_company_name], CompanyTest)

    def _page_rule_241(self, RatePages):
        # FA override — layup table and subtitle change; mechanical lift table is unchanged.
        self.compareCompanies(["FarmLayUpFactor_Ext", "MechanicalLiftFactorOtherThanZoneRated"])
        for CompanyTest in self.CompanyListDif:
            comp_name = self.extract_company_name(CompanyTest)
            self.title_company_name = CompanyTest
            if len(self.CompanyListDif) == 1:
                self.title_company_name = ""

            if self.StateAbb == "VA":
                subtitles = self._VA_RULE_SUBTITLES["Rule 241"]
            else:
                subtitles = [
                    "241.D.1.d Mechanical Lift Factor",
                    "241.E. Farm Lay-up Credit",
                ]

            if self.StateAbb == "VA":
                tables = [
                    self.buildMechanicalLiftFactor(comp_name),
                    self.buildSpecifiecCausesofLossCoverageFactor(comp_name),
                    self.buildFALayupFactors(comp_name),
                ]
            else:
                tables = [
                    self.buildMechanicalLiftFactor(comp_name),
                    self.buildFALayupFactors(comp_name),
                ]

            RatePages.generateWorksheetTablesX(
                "Rule 241 " + self.title_company_name,
                "RULE 241. PUBLIC AUTO PREMIUM DEVELOPMENT - ZONE-RATED AUTOS " + self.title_company_name,
                subtitles, tables, False, True,
            )
            self.overideFooter(RatePages.getWB()["Rule 241 " + self.title_company_name], CompanyTest)
        _wb = RatePages.getWB()
        for _sn in _wb.sheetnames:
            if _sn.startswith("Rule 241"):
                self.format241(_wb[_sn])

    def _page_rule_303(self, RatePages):
        # FA override — Rule 303 uses FarmPollutionLiability tables (different from BA).
        # Not applicable in New York.
        if self.StateAbb == "NY":
            return

        self.compareCompanies(["FarmPollutionLiabilityBaseRate_Ext",
                               "FarmPollutionLiabilityFleetFactor_Ext"])
        for CompanyTest in self.CompanyListDif:
            comp_name = self.extract_company_name(CompanyTest)
            self.title_company_name = CompanyTest
            if len(self.CompanyListDif) == 1:
                self.title_company_name = ""

            ws_name   = "Rule 303 " + self.title_company_name
            subtitles = ["303.B.1 Premium Computation", "Liability Factor:", "Fleet Factor:"]
            tables    = [
                pd.DataFrame(),                              # separator — produces the subtitle row only
                self.buildFAPollutionBaseRate(comp_name),
                self.buildFAPollutionFleetFactor(comp_name),
            ]
            RatePages.generateWorksheetTablesX(
                ws_name,
                "RULE 303. POLLUTION LIABILITY " + self.title_company_name,
                subtitles, tables, False, True,
            )
            self.overideFooter(RatePages.getWB()[ws_name], CompanyTest)

        _wb = RatePages.getWB()
        for _sn in _wb.sheetnames:
            if _sn.startswith("Rule 303"):
                self.formatRule303FA(_wb[_sn])

    def formatRule303FA(self, ws):
        from openpyxl.styles import Font, Alignment, Border, Side

        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 30

        border = Border(
            left=Side(border_style="thin", color="C1C1C1"),
            right=Side(border_style="thin", color="C1C1C1"),
            top=Side(border_style="thin", color="C1C1C1"),
            bottom=Side(border_style="thin", color="C1C1C1"),
        )
        _subtitle_vals = {"303.B.1 Premium Computation", "Liability Factor:", "Fleet Factor:"}
        _header_vals   = {"Per $25,000 Limit of Liability", "Fleet Size", "Fleet Factor"}

        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                val = str(cell.value)
                if cell.row == 1:
                    cell.font = Font(bold=True, name="Arial", size=10)
                elif val in _subtitle_vals:
                    cell.font = Font(italic=True, name="Arial", size=10)
                elif val in _header_vals:
                    cell.font      = Font(bold=True, name="Arial", size=10)
                    cell.alignment = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
                    cell.border    = border
                else:
                    cell.font      = Font(name="Arial", size=10)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border    = border

    # =========================================================================
    # Section C: Main FA page generator
    # =========================================================================

    def buildFAPages(self):
        """
        Farm Auto rate page sequence.

        This is the master list of rules for FA.
        - To ADD a new FA-only rule:   add self._page_rule_fa_xxx(RatePages) here.
        - To REMOVE a rule for FA:     comment out its line here.
        - To REORDER rules:            move its line up or down.
        - To CHANGE a rule's data:     override the buildXxx() method in Section A.

        Right now FA uses the exact same set of rules as BA, so this method
        delegates to buildBAPages(). Once FA rules diverge, replace this with
        an explicit rule list (see DEVELOPER_GUIDE.md for the full pattern).
        """
        # FA uses its own Excel factory so FA-specific generate/format
        # methods added to ExcelSettingsFA.Excel will be picked up automatically.
        from . import ExcelSettingsFA

        # Build the company list (same logic as BA)
        companies = [
            c for c in self.rateTables.keys()
            if c not in ("CW", "MM") and self.rateTables[c] is not None
        ]

        RatePages = ExcelSettingsFA.Excel(
            StateAbb=self.StateAbb,
            State=self.State,
            nEffective=self.nEffective,
            rEffective=self.rEffective,
            companyList=companies,
        )

        # Run the nesting + LCM protocol
        self.nesting()

        # Load state-specific sheet exceptions
        state_sheet_exceptions = pd.read_excel(
            BA_INPUT_FILE, sheet_name=None, engine="openpyxl"
        )

        # State-specific warnings
        if self.StateAbb == "MT":
            print("Warning: Rule 297 will not be correct.")
        if self.StateAbb == "DC":
            print("Warning: Additional PIP Base Rate Tables for 222, 232, 239 will be absent.")
        if self.StateAbb == "MI":
            print("Warning: Base Rate Formatting incomplete. 298 has special exceptions not yet built.")
        if self.StateAbb == "VA":
            print("Warning: Due to large shifts in manual presentation, this manual is incomplete.")
        if self.StateAbb in ("NY", "CA"):
            print("Warning: Rule 297 for this state was not built out.")

        # shared dict passed to rules 293, 297, and 451 (they share state)
        shared = {}

        # ── Rules ─────────────────────────────────────────────────────────────
        # Comment out any rule not needed for Farm Auto.
        # Add FA-only rules at the bottom (self._page_rule_fa_xxx).
        self._page_rule_vapcd(RatePages, state_sheet_exceptions)
        self._page_rule_208(RatePages)
        self._page_rule_222_ttt_base_rates(RatePages, state_sheet_exceptions)
        self._page_rule_222b(RatePages)
        self._page_rule_222c(RatePages)
        self._page_rule_222e(RatePages)
        self._page_rule_223b5(RatePages)
        self._page_rule_223c(RatePages)
        self._page_rule_225c2(RatePages)
        self._page_rule_225c3(RatePages)
        self._page_rule_225_zone_br(RatePages)
        self._page_rule_225d(RatePages)
        self._page_rule_231c(RatePages)
        self._page_rule_232_ppt_base_rates(RatePages, state_sheet_exceptions)
        self._page_rule_232b(RatePages)
        self._page_rule_233(RatePages)
        self._page_rule_239_school_bus_br(RatePages, state_sheet_exceptions)
        self._page_rule_239_other_bus_br(RatePages, state_sheet_exceptions)
        self._page_rule_239_van_pool_br(RatePages, state_sheet_exceptions)
        self._page_rule_239_taxi_br(RatePages, state_sheet_exceptions)
        self._page_rule_239c(RatePages)
        self._page_rule_239d(RatePages)
        self._page_rule_240(RatePages)
        self._page_rule_241(RatePages)
        self._page_rule_243(RatePages)
        self._page_rule_255(RatePages)
        self._page_rule_264(RatePages)
        self._page_rule_266(RatePages)
        self._page_rule_267(RatePages)
        self._page_rule_268(RatePages)
        self._page_rule_269(RatePages)
        self._page_rule_271(RatePages)
        self._page_rule_272(RatePages)
        self._page_rule_273(RatePages)
        self._page_rule_274(RatePages)
        self._page_rule_275(RatePages)
        self._page_rule_276(RatePages)
        self._page_rule_277(RatePages)
        self._page_rule_278(RatePages)
        self._page_rule_279(RatePages)
        self._page_rule_280(RatePages)
        self._page_rule_281(RatePages)
        self._page_rule_283(RatePages)
        self._page_rule_284(RatePages)
        self._page_rule_288(RatePages)
        self._page_rule_289(RatePages)
        self._page_rule_290(RatePages)
        self._page_rule_292(RatePages)
        self._page_rule_293(RatePages, shared)
        self._page_rule_294(RatePages)
        self._page_rule_295(RatePages)
        self._page_rule_296(RatePages)
        self._page_rule_297(RatePages, shared)
        self._page_rule_298(RatePages)
        self._page_rule_300(RatePages)
        self._page_rule_301c(RatePages)
        self._page_rule_301d1(RatePages)
        self._page_rule_301d2(RatePages)
        self._page_rule_303(RatePages)
        self._page_rule_305(RatePages)
        self._page_rule_306(RatePages)
        self._page_rule_307(RatePages)
        self._page_rule_309(RatePages)
        self._page_rule_310(RatePages)
        self._page_rule_313(RatePages)
        self._page_rule_315(RatePages)
        self._page_rule_317(RatePages)
        self._page_rule_416(RatePages)
        self._page_rule_417(RatePages)
        self._page_rule_425(RatePages)
        self._page_rule_426(RatePages)
        self._page_rule_427(RatePages)
        self._page_rule_450(RatePages)
        self._page_rule_451(RatePages, shared)
        self._page_rule_452(RatePages)
        self._page_rule_453(RatePages)
        self._page_rule_454(RatePages)
        self._page_rule_dp1(RatePages)
        self._page_rule_state_specific(RatePages)

        # ── FA-only rules (add here when Farm Auto has unique rules) ──────────
        # self._page_rule_fa_farm_machinery(RatePages)   # example

        RatePages.createIndex()

        if self.StateAbb == "FL":
            self.overideHeaderFL(RatePages.getWB())

        return RatePages.getWB()
