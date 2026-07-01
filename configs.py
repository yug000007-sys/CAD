"""
Project configurations — one entry per brand.
To add a new project: copy an existing block and fill in the mappings.
"""

PROJECTS = {

    # ─── T-SLOTS (TraceParts raw xlsx) ───────────────────────────────────────
    "T-Slots (TraceParts)": {
        "description": "TraceParts CAD download leads for Bonnell Aluminum T-Slots",
        "input_format": "xlsx",
        "source_type": "traceparts",
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
            ("Part number",               "Part number"),
            ("Part description",          "Part description"),
            ("CAD Format",                "CAD format"),
            ("Date",                      "Date"),
            ("Status",                    "Status"),
            ("Origin (Web site)",         "Origin (Web site)"),
            ("Catalog",                   "Catalog"),
            ("Opt-in (emailing allowed)", "Opt-in (emailing allowed)"),
        ],
    },

    # ─── CADENAS (msg with embedded xlsx) ────────────────────────────────────
    "Cadenas": {
        "description": "Cadenas CAD downloads — Excel attachment inside .msg",
        "input_format": "msg_xlsx",
        "source_type": "cadenas",
        "merge_by": "Email",
        "lead_intro": "This is a registered user who has downloaded the CAD drawing.",
        "lead_outro": "Please contact the customer for service and product opportunities.",
        "lead_source_1": "Website",
        "lead_source_2": "Cadenas",
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
            ("Model Number",  "CAD name"),
            ("CAD Name",      "Standardname"),
            ("Sensor Type",   "Sensor Type"),
            ("Date",          "Process end"),
        ],
    },

    # ─── NASON (passthrough — already in template format, Cadenas source) ────
    "Nason": {
        "description": "Nason leads — already in output format, merge duplicates only",
        "input_format": "xlsx",
        "source_type": "passthrough",
        "merge_by": "Email",
        "lead_intro": "This is a registered user who has downloaded the cad drawing for",
        "lead_outro": "please contact the customer for service and product opportunities.",
        "lead_source_1": "Website",
        "lead_source_2": "Cadenas",
        "col_map": {},
        "comment_fields": [],
        # Nason LeadComments pattern:
        # MODEL NUMBER:{model} and cad name: {cad_name}
        "comment_template": "nason",
    },

    # ─── WAGO (msg with embedded xlsx — PARTcommunity) ───────────────────────
    "WAGO (PARTcommunity)": {
        "description": "WAGO CAD downloads via PARTcommunity — Excel attachment in .msg",
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
            ("Part Number",    "Standardname"),
            ("CAD Name",       "CADname"),
            ("CAD Dimension",  "CADdimension"),
            ("Date",           "Process_end"),
            ("Server Type",    "Servertype"),
        ],
    },

    # ─── NEXEN (msg with CSV in body) ────────────────────────────────────────
    "Nexen Group": {
        "description": "Nexen CAD downloads — CSV data in email body (.msg)",
        "input_format": "msg_body_csv",
        "source_type": "nexen",
        "merge_by": "Email",
        "lead_intro": "This lead is generated from CAD Download:",
        "lead_outro": "Please contact the customer for product and service opportunities.",
        "lead_source_1": "Nexen",
        "lead_source_2": "CAD Downloads",
        "col_map": {
            "FirstName":     "Customer",
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
            ("Part Number",   "Product"),
            ("Content Type",  "Content Type"),
            ("Date",          "Accessed At"),
        ],
        "comment_template": "nexen",
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
            ("Industry",      "Industry"),
            ("Company Size",  "Company Size"),
        ],
    },

    # ─── WIELAND (TraceParts .msg with embedded xlsx) ─────────────────────────
    "Wieland Electric (TraceParts)": {
        "description": "Wieland Electric CAD downloads via TraceParts — Excel in .msg",
        "input_format": "msg_xlsx",
        "source_type": "traceparts",
        "merge_by": "Email",
        "lead_intro": "This is a registered user who has downloaded the CAD drawing for Wieland Electric.",
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
            ("Part number",               "Part number"),
            ("Part description",          "Part description"),
            ("Product name",              "Product name"),
            ("CAD Format",                "CAD format"),
            ("Date",                      "Date"),
            ("Status",                    "Status"),
            ("Origin (Web site)",         "Origin (Web site)"),
            ("Catalog",                   "Catalog"),
            ("Opt-in (emailing allowed)", "Opt-in (emailing allowed)"),
            ("Field of activity",         "Field of activity"),
            ("Department",                "Department/Service"),
        ],
    },

    # ─── LEAK DEFENSE / INTERLYNX (msg with CSV in body) ─────────────────────
    "Leak Defense (Interlynx)": {
        "description": "Leak Defense contractor certification leads — CSV in email body",
        "input_format": "msg_body_csv",
        "source_type": "leak_defense",
        "merge_by": "Email",
        "lead_intro": "",   # Notes to Rep field already contains the full intro
        "lead_outro": "",
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
            ("Brand",         "Brand"),
            ("Notes",         "Notes to Rep"),
        ],
        "comment_template": "leak_defense",  # Notes to Rep IS the comment body
    },

    # ─── TSUBAKI (Thomasnet xlsx) ─────────────────────────────────────────────
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
            ("Item Number",   "OrderNumber"),
            ("CAD Format",    "Format"),
            ("Date",          "DateAndTime"),
            ("Source",        "Source"),
            ("Opted In",      "OptedIn"),
            ("Download Page", "DownloadPage"),
        ],
    },

    # ─── ITT BATCH ────────────────────────────────────────────────────────────
    "ITT_Batch": {
        "description": "ITT Batch leads — grouped by Company, all contacts listed in LeadComments",
        "input_format": "auto",
        "source_type":  "itt_batch",
        "merge_by":     "Company",
        "lead_intro":   "",          # set per file in app (event name)
        "lead_outro":   "Please contact the customer for product and service opportunities.",
        "lead_source_1": "",
        "lead_source_2": "",
        "col_map": {
            "Company":   "Company",
            "Address":   "Address",   # full combined address — parsed in processor
            "_fullname": "Full Name",
            "_email":    "Email",
            "_title":    "Title",
        },
        "comment_fields": [],
        "comment_template": "itt_batch",
    },

    # ─── AMI BEARINGS (Thomasnet xlsx — passthrough) ──────────────────────────
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

    # ─── UNIVERSAL / BATCH ────────────────────────────────────────────────────
    "Universal / Batch": {
        "description": "Generic batch processor — upload any file, auto-detect columns, configure everything manually",
        "input_format": "auto",
        "source_type":  "universal",
        "merge_by":     "Email",
        "lead_intro":   "",
        "lead_outro":   "",
        "lead_source_1": "",
        "lead_source_2": "",
        "lead_source_3": "",
        "col_map":      {},
        "comment_fields": [],
    },
}
