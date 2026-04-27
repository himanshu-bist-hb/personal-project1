import openpyxl
import win32com.client
import os
import time
import gc


def process_pagebreaks(dest_filename1, dest_filename2):
    # Load the workbook
    workbook = openpyxl.load_workbook(dest_filename1)

    # Shorten sheet names if they are too long
    for sheet in workbook.sheetnames:
        if len(sheet) > 31:  # Excel sheet name limit is 31 characters
            new_name = sheet[:31]
            workbook[sheet].title = new_name

    # Iterate through all sheets in the workbook
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        sheet.print_title_rows = '1:1'
        # Re-assert page setup defaults: ExcelSettingsBA already wrote these,
        # but openpyxl can mangle pageSetUpPr/fitToPage on load+save. Resetting
        # them here guarantees they survive into the saved file.
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = False

        if sheet_name.startswith("Index"):
            sheet.print_title_rows = '0:0'  # No title rows
            sheet.print_area = 'A1:J{}'.format(sheet.max_row)  # Define print area
            sheet.page_setup.fitToWidth = 1  # Fit to page width
            sheet.page_setup.fitToHeight = False  # Disable fit to height

        elif sheet_name.startswith("Rule 222 B"):
            sheet.sheet_properties.pageSetUpPr.fitToPage = False
            sheet.page_setup.fitToWidth = False
            sheet.page_setup.fitToHeight = False
            for row in [25, 49]:
                sheet.row_breaks.append(openpyxl.worksheet.pagebreak.Break(row))

        elif sheet_name.startswith("Rule 222 TTT"):
            sheet.page_setup.fitToHeight = 1
            sheet.page_setup.fitToWidth = 1
            sheet.print_title_rows = '1:3'

        elif sheet_name.startswith("Rule 223 B.5"):
            sheet.page_setup.orientation = "landscape"

        elif sheet_name.startswith("Rule 223 C"):
            sheet.page_setup.fitToHeight = 1
            sheet.page_setup.fitToWidth = 1

        elif sheet_name.startswith("Rule 225 Zone"):
            sheet.print_area = 'A1:M{}'.format(sheet.max_row)
            sheet.print_options.horizontalCentered = False
            sheet.print_options.verticalCentered = False
            sheet.sheet_properties.pageSetUpPr.fitToPage = False
            sheet.page_setup.fitToWidth = False
            sheet.page_setup.fitToHeight = False
            for row in range(52, sheet.max_row, 51):
                sheet.row_breaks.append(openpyxl.worksheet.pagebreak.Break(row))

        elif sheet_name.startswith("Rule 225.C.3"):
            sheet.page_setup.fitToHeight = 1
            sheet.page_setup.fitToWidth = 1

        elif sheet_name.startswith("Rule 232 PPT"):
            sheet.page_setup.fitToHeight = 1
            sheet.page_setup.fitToWidth = 1
            sheet.print_title_rows = '1:3'


        elif sheet_name.startswith("Rule 239 ") and not sheet_name.startswith("Rule 239 C"):
            sheet.page_setup.fitToHeight = 1
            sheet.page_setup.fitToWidth = 1
            sheet.print_title_rows = '1:3'

        elif sheet_name.startswith("Rule 239 C"):
            sheet.page_setup.fitToHeight = 1
            sheet.page_setup.fitToWidth = 1
            sheet.page_margins.top = 1.00
            # for row in range(25, 70, 24):
            #     sheet.row_breaks.append(openpyxl.worksheet.pagebreak.Break(row))

        elif sheet_name.startswith("Rule 240 "):
            sheet.page_setup.fitToHeight = 1
            sheet.page_setup.fitToWidth = 1
            sheet.print_options.verticalCentered = True
            sheet.print_title_rows = '1:3'
            sheet.print_area = 'A1:M{}'.format(sheet.max_row)
            sheet.page_margins.top = 1.00

        elif sheet_name.startswith("Rule 255"):
            sheet.print_area = 'A1:H{}'.format(sheet.max_row)
            sheet.print_options.horizontalCentered = False
            sheet.print_options.verticalCentered = False
            sheet.sheet_properties.pageSetUpPr.fitToPage = False
            sheet.page_setup.fitToWidth = False
            sheet.page_setup.fitToHeight = False
            for row in [37]:
                sheet.row_breaks.append(openpyxl.worksheet.pagebreak.Break(row))

        elif sheet_name.startswith("Rule 275"):
            if workbook[sheet_name]["A10"].value == "275.B.1.(b). Short Term - Autos Leased by the Hour, Day, or Week":
                sheet.print_title_rows = '1:1'
                sheet.page_setup.fitToHeight = 1
                sheet.page_setup.fitToWidth = 1

        elif sheet_name.startswith("Rule 283"):
            sheet.print_area = 'A1:P{}'.format(sheet.max_row)

            target_values = ["283.B Limited Specified Causes of Loss",
                             "283.B Comprehensive",
                             "283.B Blanket Collision"]

            for row in range(1, sheet.max_row + 1):
                cell_value = str(sheet.cell(row=row, column=1).value)

                if cell_value in target_values and row > 3:
                    sheet.row_breaks.append(openpyxl.worksheet.pagebreak.Break(id=row - 1))  # Insert page break above
            sheet.page_setup.fitToWidth = 1
            sheet.page_setup.fitToHeight = False
            sheet.sheet_properties.pageSetUpPr.fitToPage = True

        elif sheet_name.startswith("Rule 289"):
            sheet.print_area = 'A1:H{}'.format(sheet.max_row)
            sheet.sheet_properties.pageSetUpPr.fitToPage = False
            sheet.page_setup.fitToWidth = False
            sheet.page_setup.fitToHeight = False
            for row in [37]:
                sheet.row_breaks.append(openpyxl.worksheet.pagebreak.Break(row))

        elif sheet_name.startswith("Rule 297"):
            sheet.print_area = 'A1:P{}'.format(sheet.max_row)
            sheet.sheet_properties.pageSetUpPr.fitToPage = False
            sheet.page_setup.fitToWidth = False
            sheet.page_setup.fitToHeight = False

            occurrence_count = 0
            for row in range(1, sheet.max_row + 1):
                cell_value = str(sheet.cell(row=row, column=1).value)
                if str(cell_value).startswith('Single') or str(cell_value).startswith('Uninsured'):
                    occurrence_count += 1
            # Add a page break before every 3rd occurrence
                if (occurrence_count % 3 == 0) and (occurrence_count != 0):
                    occurrence_count += 1
                    sheet.row_breaks.append(openpyxl.worksheet.pagebreak.Break(row-1))


        # Pretty sloppy but it works.
        elif sheet_name.startswith("Rule 298"):
            sheet.print_area = 'A1:K{}'.format(sheet.max_row)
            sheet.sheet_properties.pageSetUpPr.fitToPage = False
            sheet.page_setup.fitToWidth = False
            sheet.page_setup.fitToHeight = False

            occurrence_count = 0
            for row in range(1, sheet.max_row + 1):
                cell_value = str(sheet.cell(row=row, column=1).value)
                if str(cell_value).startswith('298'):
                    occurrence_count += 1
            # Add a page break before every 3rd occurrence
                if occurrence_count == 4: # or occurrence_count == 8:
                    occurrence_count += 1
                    sheet.row_breaks.append(openpyxl.worksheet.pagebreak.Break(row-1))
                if occurrence_count == 8:  # Dont need more breaks at this point
                    break

        # VA has custom 301 names that conflict with normal logic. added if else statements to account for this
        elif sheet_name.startswith(("Rule 301.A", "Rule 301.B", "Rule 301.C")):
            if workbook[sheet_name]["B4"].value in ["Extra Heavy Truck-Tractor", "Extra-Heavy Truck", "Heavy Truck",
                                                    "Heavy Truck-Tractor", "Light Truck", "Medium Truck", "Private Passenger Types",
                                                    "Semitrailer", "Service or Utility Trailer", "Trailer"]:
                pass
            else:
                sheet.page_setup.fitToWidth = 1
                sheet.page_setup.fitToHeight = False
                sheet.sheet_properties.pageSetUpPr.fitToPage = True
                if sheet_name.startswith("Rule 301.B"): # sloppy solution to bypass automatic page breaks. Column T is minimum width to solve this way
                    sheet.print_area = 'A1:T{}'.format(sheet.max_row)
                for row in range(46, sheet.max_row, 45):
                    sheet.row_breaks.append(openpyxl.worksheet.pagebreak.Break(row))
                sheet.print_options.horizontalCentered = False
                sheet.print_options.verticalCentered = False
                sheet.page_setup.orientation = "landscape"
                sheet.page_margins.top = 1.00

        elif sheet_name.startswith(("Rule 301.C", "Rule 301.D)")):
            if workbook[sheet_name]["B4"].value in ["Extra Heavy Truck-Tractor", "Extra-Heavy Truck", "Heavy Truck",
                                                    "Heavy Truck-Tractor", "Light Truck", "Medium Truck", "Private Passenger Types",
                                                    "Semitrailer", "Service or Utility Trailer", "Trailer"]:
                # This margin collides with FLs special stamping.
                if not "FL" in dest_filename1:
                    sheet.page_margins.top = 1.00

                sheet.page_setup.fitToWidth = 1
                sheet.page_setup.fitToHeight = 1
            else:
                pass

        elif sheet_name.startswith("Rule 306"):
            sheet.page_setup.fitToWidth = 1
            sheet.print_title_rows = '1:4'

        elif sheet_name.startswith("Rule 315"):
            sheet.page_setup.fitToWidth = 1
            sheet.page_setup.fitToHeight = False
            sheet.sheet_properties.pageSetUpPr.fitToPage = True
            for row in [23]:
                sheet.row_breaks.append(openpyxl.worksheet.pagebreak.Break(row))

        # Pretty sloppy but it works.
        elif sheet_name.startswith("Rule R1"):
            sheet.print_area = 'A1:M{}'.format(sheet.max_row)
            sheet.sheet_properties.pageSetUpPr.fitToPage = False
            sheet.page_setup.fitToWidth = False
            sheet.page_setup.fitToHeight = False

            occurrence_count = 0
            for row in range(1, sheet.max_row + 1):
                cell_value = str(sheet.cell(row=row, column=1).value)
                if str(cell_value).startswith('R1'):
                    occurrence_count += 1
            # Add a page break before every 3rd occurrence
                if occurrence_count == 3 or occurrence_count == 6: # or occurrence_count == 8:
                    occurrence_count += 1
                    sheet.row_breaks.append(openpyxl.worksheet.pagebreak.Break(row-1))

    # Save the modified workbook with a new name
    # modified_filename = dest_filename1.replace('.xlsx', '_modified.xlsx')
    workbook.save(dest_filename1)
    workbook.close() # Close the workbook at the end
    print("Stage 3: Page Breaks saved.")

    time.sleep(2)

    # Normalize paths
    dest_filename1 = os.path.normpath(os.path.abspath(dest_filename1))
    dest_filename2 = os.path.normpath(os.path.abspath(dest_filename2))
    repaired_filename = dest_filename1.replace(".xlsx", "_repaired.xlsx")

    # Ensure Excel is not running
    def kill_excel_instances():
        import subprocess
        subprocess.call("taskkill /f /im excel.exe", shell=True)

    kill_excel_instances()
    time.sleep(2)

    xlApp = win32com.client.DispatchEx('Excel.Application')
    xlApp.Visible = False

    try:
        # Step 1: Open with CorruptLoad to repair
        xlBook = xlApp.Workbooks.Open(dest_filename1, CorruptLoad=1)
        xlBook.SaveAs(repaired_filename)
        xlBook.Close(False)
        print("File repaired and saved.")
    except Exception as e:
        print(f"Error during repair: {e}")
        xlApp.Quit()
        return

    # Step 2: Reopen the repaired file normally
    try:
        xlBook = xlApp.Workbooks.Open(repaired_filename, ReadOnly=True)
        xlApp.Visible = True

        index_sheet = xlBook.Sheets("Index")
        index_sheet.Visible = False

        # time.sleep(3)
        # xlBook.ExportAsFixedFormat(0, dest_filename2, Quality=0)
        # print("PDF export complete.")
    except Exception as e:
        print(f"Error exporting to PDF: {e}")
    finally:
        xlBook.Close(False)
        xlApp.Quit()

        # Clean up COM objects
        del xlBook
        del xlApp
        gc.collect()

        # if os.path.exists(dest_filename2):
        #     print(f"PDF file created successfully: {dest_filename2}")
        # else:
        #     print("PDF creation failed.")

        os.remove(dest_filename1)
        print("Corrupted File deleted successfully.")

        os.rename(repaired_filename, dest_filename1)

        print("Waiting 10 seconds for items to flush to disk.")
        time.sleep(10)


# Example usage
# process_pagebreaks(r'C:\Users\bernb17\Nationwide\Desktop\CL-State-Pages-Dump\ME - ISO Curr', 'ME 03-01-25 BA Small Market Rate Pages.xlsx', 'output.pdf')
