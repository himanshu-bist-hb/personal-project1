"""
One-time script to create FA/FA_Rule223C2_ClassMapping.xlsx.
Run from the project root: python FA/create_223c2_mapping.py
"""
import pandas as pd
from pathlib import Path

# Each row: [Primary Class, Secondary Class display, Class Code (2-digit), Farmers Use match key]
ROWS = [
    # ── Truckers ────────────────────────────────────────────────────────────────
    ["Truckers", "Carrier Both Private Carriage & Transport\nGoods/Materials/Commodities", "02", "Not Applicable"],
    ["Truckers", "Tow Trucks For-Hire",                                                    "03", "Not Applicable"],
    ["Truckers", "Common Carriers",                                                         "21", "Not Applicable"],
    ["Truckers", "Contract Carriers (Other than Chemical or Iron and\nSteel Haulers)",     "22", "Not Applicable"],
    ["Truckers", "Contract Carriers Hauling Chemicals",                                    "23", "Not Applicable"],
    ["Truckers", "Contract Carriers Hauling Iron and Steel",                               "24", "Not Applicable"],
    ["Truckers", "Exempt Carriers (Other than Livestock Haulers)",                         "25", "Not Applicable"],
    ["Truckers", "Exempt Carriers Hauling Livestock",                                      "26", "Not Applicable"],
    ["Truckers", "All Other",                                                              "29", "Not Applicable"],
    # ── Food Delivery ────────────────────────────────────────────────────────────
    ["Food Delivery", "Canneries and Packing Plants",  "31", "Not Applicable"],
    ["Food Delivery", "Fish and Seafood",              "32", "Not Applicable"],
    ["Food Delivery", "Frozen Food",                   "33", "Not Applicable"],
    ["Food Delivery", "Fruit and Vegetable",           "34", "Not Applicable"],
    ["Food Delivery", "Meat or Poultry",               "35", "Not Applicable"],
    ["Food Delivery", "All Other",                     "39", "Not Applicable"],
    # ── Specialized Delivery ─────────────────────────────────────────────────────
    ["Specialized Delivery", "Armored Cars",             "41", "Not Applicable"],
    ["Specialized Delivery", "Film Delivery",            "42", "Not Applicable"],
    ["Specialized Delivery", "Magazines or Newspapers",  "43", "Not Applicable"],
    ["Specialized Delivery", "Mail and Parcel Post",     "44", "Not Applicable"],
    ["Specialized Delivery", "All Other",                "49", "Not Applicable"],
    # ── Waste Disposal ───────────────────────────────────────────────────────────
    ["Waste Disposal", "Auto Dismantlers",          "51", "Not Applicable"],
    ["Waste Disposal", "Building Wrecking Operators","52", "Not Applicable"],
    ["Waste Disposal", "Garbage",                   "53", "Not Applicable"],
    ["Waste Disposal", "Junk Dealers",              "54", "Not Applicable"],
    ["Waste Disposal", "All Other",                 "59", "Not Applicable"],
    # ── Farmers – code 61 ────────────────────────────────────────────────────────
    ["Farmers", "Hobby Farm",                                                                       "61", "Hobby Farm"],
    ["Farmers", "Hauling Chemicals or Petroleum",                                                   "61", "Hauling Chemicals or Petroleum"],
    ["Farmers", "Special Use Vehicles – More than Incidental Use",                             "61", "Special Use Vehicles - More than Incidental Use"],
    ["Farmers", "Dump Vehicles - Including Hauling for Hire",                                       "61", "Dump Vehicles - Including Hauling for Hire"],
    ["Farmers", "Dump Vehicles - No Hauling for Hire",                                              "61", "Dump Vehicles - No Hauling for Hire"],
    ["Farmers", "Vehicles Used to Transport Farm Commodities -\nIncluding Hauling for Hire",        "61", "Vehicles Used to Transport Farm Commodities - Including Hauling for Hire"],
    ["Farmers", "Vehicles Used to Transport Farm Commodities - No\nHauling for Hire",               "61", "Vehicles Used to Transport Farm Commodities - No Hauling for Hire"],
    ["Farmers", "Vehicles Used to Transport Milk and Similar Liquid\nCargo - Including Hauling for Hire", "61", "Vehicles Used to Transport Milk and Similar Liquid Cargo - Including Hauling for Hire"],
    ["Farmers", "Vehicles Used to Transport Milk and Similar Liquid\nCargo - No Hauling for Hire",  "61", "Vehicles Used to Transport Milk and Similar Liquid Cargo - No Hauling for Hire"],
    ["Farmers", "All Other",                                                                        "61", "All Other"],
    # ── Farmers – code 62 ────────────────────────────────────────────────────────
    ["Farmers", "Livestock – Including Hauling for Hire", "62", "Livestock - Including Hauling for Hire"],
    ["Farmers", "Livestock – No Hauling for Hire",        "62", "Livestock - No Hauling for Hire"],
    # ── Farmers – code 69 ────────────────────────────────────────────────────────
    ["Farmers", "Hauling Chemicals or Petroleum",                                                   "69", "Hauling Chemicals or Petroleum"],
    ["Farmers", "Special Use Vehicles – More than Incidental Use",                             "69", "Special Use Vehicles - More than Incidental Use"],
    ["Farmers", "Dump Vehicles - Including Hauling for Hire",                                       "69", "Dump Vehicles - Including Hauling for Hire"],
    ["Farmers", "Dump Vehicles - No Hauling for Hire",                                              "69", "Dump Vehicles - No Hauling for Hire"],
    ["Farmers", "Vehicles Used to Transport Farm Commodities -\nIncluding Hauling for Hire",        "69", "Vehicles Used to Transport Farm Commodities - Including Hauling for Hire"],
    ["Farmers", "Vehicles Used to Transport Farm Commodities - No\nHauling for Hire",               "69", "Vehicles Used to Transport Farm Commodities - No Hauling for Hire"],
    ["Farmers", "Vehicles Used to Transport Milk and Similar Liquid\nCargo - Including Hauling for Hire", "69", "Vehicles Used to Transport Milk and Similar Liquid Cargo - Including Hauling for Hire"],
    ["Farmers", "Vehicles Used to Transport Milk and Similar Liquid\nCargo - No Hauling for Hire",  "69", "Vehicles Used to Transport Milk and Similar Liquid Cargo - No Hauling for Hire"],
    ["Farmers", "All Other",                                                                        "69", "All Other"],
    # ── Dump and Transit Mix ─────────────────────────────────────────────────────
    ["Dump and Transit Mix", "Excavating",                         "71", "Not Applicable"],
    ["Dump and Transit Mix", "Sand and Gravel (Other than Quarrying)", "72", "Not Applicable"],
    ["Dump and Transit Mix", "Mining",                             "73", "Not Applicable"],
    ["Dump and Transit Mix", "Quarrying",                          "74", "Not Applicable"],
    ["Dump and Transit Mix", "All Other",                          "79", "Not Applicable"],
    # ── Contractors ──────────────────────────────────────────────────────────────
    ["Contractors", "Building – Commercial",                                          "81", "Not Applicable"],
    ["Contractors", "Building – Private Dwellings",                                   "82", "Not Applicable"],
    ["Contractors", "Electrical, Plumbing, Masonry, Plastering, Other\nRepair/Service",   "83", "Not Applicable"],
    ["Contractors", "Excavating",                                                          "84", "Not Applicable"],
    ["Contractors", "Street and Road",                                                     "85", "Not Applicable"],
    ["Contractors", "All Other",                                                           "89", "Not Applicable"],
    # ── Not Otherwise Specified ──────────────────────────────────────────────────
    ["Not Otherwise Specified", "Logging and Lumbering", "91", "Not Applicable"],
    ["Not Otherwise Specified", "All Other",             "99", "Not Applicable"],
]

df = pd.DataFrame(ROWS, columns=["Primary Class", "Secondary Class", "Class Code", "Farmers Use"])

out = Path(__file__).parent / "FA_Rule223C2_ClassMapping.xlsx"
df.to_excel(out, sheet_name="ClassMapping", index=False)
print(f"Created: {out}")
