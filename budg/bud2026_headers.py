# ============================
# BUD2026 QUARTERLY MODEL - column wiring
# ============================
#
# The output layout (headers, formulas, styling) lives in bud2026_template.json,
# extracted verbatim from "AR Collection and Provision Forecast - Master
# quarterly.xlsx". This module only maps the mapper's friendly column names to
# the ALL-sheet column letters the export writes values into.

# bud_rows column name -> ALL-sheet column letter (static values written by the tool)
VALUE_COLUMNS = {
    "CustCode": "A",
    "Cust Name": "B",
    "BT": "C",
    "Sales Budget region": "D",
    "Cust Region": "E",
    "Customer Status": "F",
    "Main Ac": "G",
    "Insurance": "H",
    "On\nAccount": "I",
    "Not Due\nAmount": "J",
    "Not Due\n0-30 days": "K",
    "Not Due\n31-60 days": "L",
    "Not Due\n61-90 days": "M",
    "Not Due\n91-180 days": "N",
    "Not Due\n180+ days": "O",
    "Aging\n1 to 30": "P",
    "Aging\n31 to 60": "Q",
    "Aging\n61 to 90": "R",
    "Aging\n91 to 120": "S",
    "Aging\n121 to 150": "T",
    "Aging\n>=151": "U",
    " AR\nBalance": "V",
}

# Collections FC (FIFO) input columns, pre-filled from By_Customer when a
# source column exists (blank otherwise; blank and 0 behave identically).
COLLECTION_FC_COLUMNS = {
    "Collections FC\n31/03/2026": "AD",
    "Collections FC\n30/06/2026": "AS",
    "Collections FC\n30/09/2026": "BH",
    "Collections FC\n31/12/2026": "BW",
}

QUARTER_COLLECTION_HEADERS = {
    "Q1": "Collections FC\n31/03/2026",
    "Q2": "Collections FC\n30/06/2026",
    "Q3": "Collections FC\n30/09/2026",
    "Q4": "Collections FC\n31/12/2026",
}

# FY2026 quarter end dates (month, day) in order
QUARTER_ENDS_2026 = [(3, 31), (6, 30), (9, 30), (12, 31)]
