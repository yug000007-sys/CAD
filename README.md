# Leads Agent

Multi-project CAD lead processor. Upload a raw file (`.xlsx` or `.msg`), select your project, and download a clean CSV ready for your CRM.

## Projects supported

| Project | Input | Source |
|---|---|---|
| T-Slots (TraceParts) | `.xlsx` | TraceParts raw export |
| Cadenas | `.msg` (Excel attachment) | Cadenas portal |
| WAGO (PARTcommunity) | `.msg` (Excel attachment) | PARTcommunity/WAGO |
| Nexen Group | `.msg` (CSV in body) | Nexen portal |
| PROCO Products | `.msg` (Excel attachment) | TraceParts/Proco |
| Leak Defense (Interlynx) | `.msg` (CSV in body) | Interlynx/Watts |
| Nason | `.xlsx` | Already-processed file, deduplication only |
| Tsubaki | `.xlsx` | Thomasnet CAD downloads |
| AMI Bearings | `.xlsx` | Already-processed file, deduplication only |

## Features

- ✅ Merges all rows by email address
- ✅ Builds `LeadComments` in HTML format per download entry
- ✅ Auto-detects and translates non-Latin text (Korean, Arabic, Chinese, etc.) via Claude AI
- ✅ Handles `.msg` files with embedded Excel or CSV-in-body
- ✅ Alphanumeric zip codes (UK, Canada, etc.) preserved as-is
- ✅ 51-column output template compatible with your CRM

## Setup

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/leads-agent.git
cd leads-agent

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run locally
streamlit run app.py
```

## Deploy to Streamlit Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo
4. Set main file path: `app.py`
5. Add your Anthropic API key as a secret (optional — only needed for non-Latin translation):
   - In Streamlit Cloud → App settings → Secrets
   - Add: `ANTHROPIC_API_KEY = "sk-ant-..."`

## Adding a new project

Open `projects/configs.py` and copy any existing block. Fill in:

```python
"Your Project Name": {
    "description": "What this project does",
    "input_format": "xlsx",          # xlsx | msg_xlsx | msg_body_csv
    "source_type": "your_project",
    "merge_by": "Email",
    "lead_intro": "Intro text for LeadComments...",
    "lead_outro": "Outro text for LeadComments...",
    "lead_source_1": "Source Name",
    "lead_source_2": "CAD Downloads",
    "col_map": {
        "FirstName":     "First name column in raw file",
        "LastName":      "Last name column",
        "ContactTitle":  "Job title column",
        "Company":       "Company column",
        "Address":       ["Address1", "Address2"],   # list for multi-part address
        "City":          "City column",
        "State":         "State column",
        "ZipCode":       "Zip column",
        "Country":       "Country column",
        "PhoneSupplied": "Phone column",
    },
    "comment_fields": [
        ("Label in HTML",  "raw_column_name"),
        ...
    ],
},
```

That's it — no other code changes needed.

## File structure

```
leads-agent/
├── app.py                  # Streamlit UI
├── requirements.txt
├── README.md
├── projects/
│   ├── __init__.py
│   └── configs.py          # All project configs — add new projects here
└── utils/
    ├── __init__.py
    └── processor.py        # Core processing engine
```
