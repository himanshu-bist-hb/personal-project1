"""
FA/ExcelSettingsFA.py
=====================
Excel workbook factory for Farm Auto rate pages.

FA currently uses the IDENTICAL layout, formatting, and generate methods as BA.
This file simply re-exports BA's Excel class so no code is duplicated.

HOW TO ADD FA-SPECIFIC FORMATTING
----------------------------------
If FA ever needs a different layout for a specific sheet, add a method here:

    from BA.ExcelSettingsBA import Excel as _BABase

    class Excel(_BABase):

        def generateWorksheetFACustom(self, sheet_name, title, subtitle, df, ...):
            # FA-only layout logic here
            ...

        def formatFACustom(self, ws):
            # FA-only formatting here
            ...

Until then, `class Excel(_BABase): pass` is all that is needed.
"""

from BA.ExcelSettingsBA import Excel as _BABase


class Excel(_BABase):
    """
    Farm Auto Excel factory.
    Inherits ALL generate/format methods from BA.
    Add FA-specific methods here only when FA genuinely needs different layouts.
    """
    pass
