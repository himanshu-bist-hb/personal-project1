# This module builds the Habitational (Hab) State Page workbook — Appetite
# version, layered on top of EITHER base version (BP-2.0 or pre-2.0).
#
# _HabAppetiteMixin defines the Appetite-only table ("H Table 4.C. Exceptions
# to Habitational - Premium Development") once; Hab/HabCurrent below mix it
# onto BP-2.0's HabPage.Hab / pre-2.0's HabPageCurrent.Hab respectively, so
# neither base's own build*()/format*() methods are transcribed twice.

from .HabPage import Hab as HabBP20
from .HabPageCurrent import Hab as HabPre20


class _HabAppetiteMixin:
    # Builds the "H Table 4.C. Exceptions to Habitational - Premium
    # Development" table (Appetite-only, every state): the single
    # HabitabilityExclusionFactor rate filed for the "allperil" peril in
    # "BP7_Peril ExclusionOfHabitabilityClaims_Factor".
    # Returns a dataframe
    def buildHabExclusionPremiumDev(self):
        habExclusionPremiumDev = self.buildDataFrame("BP7_Peril ExclusionOfHabitabilityClaims_Factor")
        return habExclusionPremiumDev[habExclusionPremiumDev['Peril TypeCodeCode'] == 'allperil']. \
                rename(columns={'HabitabilityExclusionFactor': 'Rate'}).filter(items=['Rate'])

    # Extends the base version's sheet list with the Appetite-only page,
    # inserted at the position matching its rule number. "HPD" (4.C) sorts
    # right after "PLUS" (4.B) and ahead of the CA-only "HABEX" (also 4.C,
    # kept unchanged from BP-2.0) since 4.C. Exceptions precedes 4.C
    # Habitability Exclusion within the same rule number. On pre-2.0 (which
    # has no HABEX at all), the `next(..., len(sheetSpecs))` fallback simply
    # appends HPD at the end.
    def _sheetSpecs(self):
        sheetSpecs = super()._sheetSpecs()
        if not any(spec[0] == 'HPD' for spec in sheetSpecs):
            insertAt = next((i for i, spec in enumerate(sheetSpecs) if spec[0] == 'HABEX'), len(sheetSpecs))
            sheetSpecs.insert(insertAt, ('HPD', 'H Table 4.C. Exceptions to Habitational - Premium Development',
                                          self.buildHabExclusionPremiumDev, False, True, None, None))
        return sheetSpecs


class Hab(_HabAppetiteMixin, HabBP20):
    pass


class HabCurrent(_HabAppetiteMixin, HabPre20):
    pass
