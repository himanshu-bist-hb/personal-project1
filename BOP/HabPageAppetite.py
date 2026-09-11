# This module builds the Habitational (Hab) State Page workbook — Appetite
# version.
#
# Appetite is built directly on top of BP-2.0: this class subclasses
# HabPage.Hab and only defines what's new (the "Appetite" pages). Every
# other table — build*() method, layout, formatting — is inherited unchanged
# from HabPage.Hab, so BP-2.0 stays the single source of truth and nothing
# is transcribed twice. New Appetite pages get appended to the sheet list
# built by Hab._sheetSpecs(), inserted at the position matching their rule
# number (see the base class for the existing list this extends).

from .HabPage import Hab as HabBP20


class Hab(HabBP20):
    # Builds the "H Table 4.C. Exceptions to Habitatational - Premium
    # Development" table (Appetite-only, every state): the single
    # HabitabilityExclusionFactor rate filed for the "allperil" peril in
    # "BP7_PerilExclusionOfHabitabilityClaims_Factor".
    # Returns a dataframe
    def buildHabExclusionPremiumDev(self):
        habExclusionPremiumDev = self.buildDataFrame("BP7_PerilExclusionOfHabitabilityClaims_Factor")
        return habExclusionPremiumDev.query('`Peril TypeCode` == "allperil"'). \
                rename(columns={'HabitabilityExclusionFactor': 'Rate'}).filter(items=['Rate'])

    # Extends BP-2.0's sheet list with the Appetite-only pages, inserted at
    # the position matching their rule number. "HABEXPD" (4.C) sorts right
    # after "PLUS" (4.B) and ahead of the CA-only "HABEX" (also 4.C, kept
    # unchanged from BP-2.0) since 4.C. Exceptions precedes 4.C Habitability
    # Exclusion within the same rule number.
    def _sheetSpecs(self):
        sheetSpecs = super()._sheetSpecs()
        insertAt = next((i for i, spec in enumerate(sheetSpecs) if spec[0] == 'HABEX'), len(sheetSpecs))
        sheetSpecs.insert(insertAt, ('HABEXPD', 'H Table 4.C. Exceptions to Habitatational - Premium Development',
                                      self.buildHabExclusionPremiumDev, False, True, None, None))
        return sheetSpecs
