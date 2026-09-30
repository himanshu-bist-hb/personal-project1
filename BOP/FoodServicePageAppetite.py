# This module builds the Food Service State Page workbook — Appetite
# version, layered on top of EITHER base version (BP-2.0 or pre-2.0).
#
# Appetite tables are always sourced from BP-2.0 (FoodServicePage.py), never
# pre-2.0, per the user's instruction. _FoodAppetiteMixin holds the one
# BP-2.0-sourced method this program needs (buildFSSplzdEndo) plus a
# _sheetSpecs() override that's idempotent (safe to mix onto either base):
#
#   - FS Table 4.A.1: pre-2.0's own sheet list has "PLUS" (Food Service PLUS
#     Endorsement) at this same table number — a DIFFERENT table than
#     BP-2.0's "FSS" (Food Service Specialized Endorsement). Per the user's
#     rule ("if pre2.0 has a table with some number and that table is also
#     in the appetite, provide the appetite table code instead"), PLUS is
#     removed and FSS is inserted in its place. On a BP-2.0 base, FSS is
#     already there (pre-2.0's PLUS never was), so the removal is a no-op
#     and the insertion is skipped since FSS already exists — same
#     mixin works unmodified on both bases.

import pandas as pd

from .FoodServicePage import Food as FoodBP20
from .FoodServicePageCurrent import Food as FoodPre20


class _FoodAppetiteMixin:
    # FS Table 4.A.1. Food Service Specialized Endorsement — copied verbatim
    # from FoodServicePage.Food.buildFSSplzdEndo (BP-2.0).
    def buildFSSplzdEndo(self):
        endorsementCharge = self.buildDataFrame("BP7_MiscellaneousSpecializedEndorsement_Charges")
        rows = endorsementCharge[endorsementCharge['SpecializedEndorsementName'] == 'Food Service Specialized Endorsement']
        charge = float(rows['EndorsementCharge'].iloc[0])
        return pd.DataFrame({"Base premium for each Food Service Premises": ["${0:,.2f}".format(charge)]})

    def _sheetSpecs(self):
        sheetSpecs = super()._sheetSpecs()
        sheetSpecs = [spec for spec in sheetSpecs if spec[0] != 'PLUS']
        if not any(spec[0] == 'FSS' for spec in sheetSpecs):
            insertAt = next((i for i, spec in enumerate(sheetSpecs) if spec[0] == 'SPO'), len(sheetSpecs))
            sheetSpecs.insert(insertAt, ('FSS', 'FS Table 4.A.1. Food Service Specialized Endorsement', self.buildFSSplzdEndo, False, True, 'PLUS', None))
        return sheetSpecs


class Food(_FoodAppetiteMixin, FoodBP20):
    pass


class FoodCurrent(_FoodAppetiteMixin, FoodPre20):
    pass
