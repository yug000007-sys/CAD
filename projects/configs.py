"""
Project configurations — one entry per brand.
To add a new project: copy an existing block and fill in the mappings.
"""

PROJECTS = {

    # ─── T-SLOTS (TraceParts raw xlsx) ───────────────────────────────────────
    "T-Slots (TraceParts)": {
        "description": "TraceParts CAD download leads for Bonnell Aluminum T-Slots",
        "input_format": "xlsx",          # xlsx | msg_xlsx | msg_csv | msg_body_csv
        "source_type": "traceparts",     # used by processor to pick parsing logic
        "merge_by": "Email",
        "lead_intro": "This is a registered user who has downloaded the CAD drawing for T SLOTS.",
        "lead_outro": "Please contact the customer for services or product opportunities.",
        "lead_source_1": "Traceparts",
        "lead_source_2": "CAD Downloads",
        "col_map": {
            "FirstName":     "First name",
            "LastName":      "Name",
            "ContactTitle":  "Job",
            "Company":       "Company",
            "Address":       ["Address1", "Address2", "Address3"],
            "City":          "City",
            "State":         "State (USA)",
            "ZipCode":       "Zip/Postal Code",
            "Country":       "Country name",
            "PhoneSupplied": "Phone",
        },
        "comment_fields": [
            ("Part number",              "Part number"),
            ("Part description",         "Part description"),
            ("CAD Format",               "CAD format"),
            ("Date",                     "Date"),
            ("Status",                   "Status"),
            ("Origin (Web site)",        "Origin (Web site)"),
            ("Catalog",                  "Catalog"),
            ("Opt-in (emailing allowed)","Opt-in (emailing allowed)"),
        ],
    },

    # ─── CADENAS (msg with embedded xlsx, no header row) ─────────────────────
    "Cadenas": {
        "description": "Cadenas CAD downloads — Excel attachment inside .msg",
        "input_format": "msg_xlsx",
        "source_type": "cadenas",
        "msg_header_row": 0,            # row index that contains headers (0 = first row IS header)
        "merge_by": "Email",
        "lead_intro": "This is a registered user who has downloaded the CAD drawing via Cadenas.",
        "lead_outro": "Please contact the customer for services or product opportunities.",
        "lead_source_1": "Cadenas",
        "lead_source_2": "CAD Downloads",
        "col_map": {
            "FirstName":     "Name",
            "LastName":      "",
            "ContactTitle":  "",
            "Company":       "Firm",
            "Address":       ["Street"],
            "City":          "City",
            "State":         "State",
            "ZipCode":       "Zip",
            "Country":       "Country",
            "PhoneSupplied": "Phone",
        },
        "comment_fields": [
            ("CAD Name",    "CAD name"),
            ("Standard",    "Standardname"),
            ("Date",        "Process end"),
        ],
    },

    # ─── WAGO (msg with embedded xlsx — WEEK_24 style) ───────────────────────
    "WAGO (PARTcommunity)": {
        "description": "WAGO CAD downloads via PARTcommunity — Excel in .msg",
        "input_format": "msg_xlsx",
        "source_type": "wago",
        "merge_by": "Email",
        "lead_intro": "This is a registered user who has downloaded the CAD drawing for WAGO.",
        "lead_outro": "Please contact the customer for services or product opportunities.",
        "lead_source_1": "PARTcommunity",
        "lead_source_2": "CAD Downloads",
        "col_map": {
            "FirstName":     "Name",
            "LastName":      "",
            "ContactTitle":  "",
            "Company":       "Firm",
            "Address":       ["Street"],
            "City":          "City",
            "State":         "State",
            "ZipCode":       "Zip",
            "Country":       "Country",
            "PhoneSupplied": "Phone",
        },
        "comment_fields": [
            ("CAD Name",         "CADname"),
            ("Standard",         "Standardname"),
            ("Dimension",        "CADdimension"),
            ("Date",             "Process_end"),
            ("Server Type",      "Servertype"),
        ],
    },

    # ─── NEXEN (msg with CSV in body — Downloads_6_17 style) ─────────────────
    "Nexen Group": {
        "description": "Nexen CAD downloads — CSV data in email body",
        "input_format": "msg_body_csv",
        "source_type": "nexen",
        "merge_by": "Email",
        "lead_intro": "This is a registered user who has downloaded the CAD drawing for Nexen Group.",
        "lead_outro": "Please contact the customer for product and service opportunities.",
        "lead_source_1": "Nexen",
        "lead_source_2": "CAD Downloads",
        "col_map": {
            "FirstName":     "Customer",     # "Alex Best (12282)" → strip ID
            "LastName":      "",
            "ContactTitle":  "",
            "Company":       "Company",
            "Address":       ["Address 1", "Address 2"],
            "City":          "City",
            "State":         "State/Province",
            "ZipCode":       "Zip",
            "Country":       "Country",
            "PhoneSupplied": "Phone",
        },
        "comment_fields": [
            ("Product",       "Product"),
            ("Content Type",  "Content Type"),
            ("Date",          "Accessed At"),
        ],
    },

    # ─── PROCO (msg with embedded xlsx — FW Daily Catalog Statistics) ─────────
    "PROCO Products": {
        "description": "PROCO Products CAD downloads via TraceParts — Excel in .msg",
        "input_format": "msg_xlsx",
        "source_type": "proco",
        "merge_by": "Email",
        "lead_intro": "This is a registered user who has downloaded the CAD drawing for PROCO Products.",
        "lead_outro": "Please contact the customer for services or product opportunities.",
        "lead_source_1": "ProcoProducts.com",
        "lead_source_2": "CAD Downloads",
        "col_map": {
            "FirstName":     "First Name",
            "LastName":      "Last Name",
            "ContactTitle":  "Title",
            "Company":       "Company",
            "Address":       ["Address1", "Address2", "Address3"],
            "City":          "City",
            "State":         "State",
            "ZipCode":       "Postal Code",
            "Country":       "Country",
            "PhoneSupplied": "Phone",
        },
        "comment_fields": [
            ("Part Number",   "Part Number"),
            ("Description",   "Description"),
            ("CAD Format",    "CAD Format"),
            ("Date",          "Download Date"),
            ("Origin",        "Origin of Download"),
            ("Catalog",       "Catalog"),
            ("Opt-In",        "Opt In Status"),
        ],
    },

    # ─── LEAK DEFENSE / INTERLYNX (msg with CSV in body) ─────────────────────
    "Leak Defense (Interlynx)": {
        "description": "Leak Defense contractor certification leads — CSV in email body",
        "input_format": "msg_body_csv",
        "source_type": "leak_defense",
        "merge_by": "Email",
        "lead_intro": "This person recently completed the Leak Defense Installer Certification program.",
        "lead_outro": "Please review their qualifications and follow-up with them to address any questions.",
        "lead_source_1": "Social-Facebook",
        "lead_source_2": "Become a Leak Defense Certified Contractor",
        "col_map": {
            "FirstName":     "First Name",
            "LastName":      "Last Name",
            "ContactTitle":  "Persona",
            "Company":       "Company Name",
            "Address":       [],
            "City":          "City",
            "State":         "State",
            "ZipCode":       "Postal Code",
            "Country":       "Country",
            "PhoneSupplied": "",
        },
        "comment_fields": [
            ("Lead Source 3", "Lead Source 3"),
            ("Notes",         "Notes to Rep"),
        ],
    },

    # ─── NASON (already-processed main file, pass-through) ───────────────────
    "Nason": {
        "description": "Nason leads — already in output format, merge duplicates only",
        "input_format": "xlsx",
        "source_type": "passthrough",   # columns already match template
        "merge_by": "Email",
        "lead_intro": "",
        "lead_outro": "",
        "lead_source_1": "Website",
        "lead_source_2": "Cadenas",
        "col_map": {},                  # passthrough — columns already named correctly
        "comment_fields": [],
    },

    # ─── TSUBAKI (xlsx with different schema) ────────────────────────────────
    "Tsubaki": {
        "description": "Tsubaki CAD user downloads from Thomasnet",
        "input_format": "xlsx",
        "source_type": "tsubaki",
        "merge_by": "Email",
        "lead_intro": "This is a registered user who has downloaded the CAD drawing for Tsubaki.",
        "lead_outro": "Please contact the customer for product and service opportunities.",
        "lead_source_1": "Thomasnet",
        "lead_source_2": "CAD Downloads",
        "col_map": {
            "FirstName":     "FirstName",
            "LastName":      "LastName",
            "ContactTitle":  "",
            "Company":       "Company",
            "Address":       [],
            "City":          "",
            "State":         "",
            "ZipCode":       "Zip",
            "Country":       "Country",
            "PhoneSupplied": "",
        },
        "comment_fields": [
            ("Product",       "Breadcrumb_P1"),
            ("Order Number",  "OrderNumber"),
            ("Format",        "Format"),
            ("Date",          "DateAndTime"),
            ("Source",        "Source"),
        ],
    },

    # ─── AMI BEARINGS (already-processed main file) ───────────────────────────
    "AMI Bearings": {
        "description": "AMI Bearings leads — already in output format, merge duplicates only",
        "input_format": "xlsx",
        "source_type": "passthrough",
        "merge_by": "Email",
        "lead_intro": "",
        "lead_outro": "",
        "lead_source_1": "Website",
        "lead_source_2": "CAD Download",
        "col_map": {},
        "comment_fields": [],
    },
}
