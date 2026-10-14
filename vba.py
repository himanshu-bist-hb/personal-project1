from oletools.olevba import VBA_Parser

file_path = r"C:\Users\YourName\Documents\input.xlsm"

vba = VBA_Parser(file_path)

if vba.detect_vba_macros():
    for filename, stream_path, vba_filename, code in vba.extract_macros():
        print(f"\n--- {vba_filename} ---\n")
        print(code)
else:
    print("No VBA macros found.")

vba.close()