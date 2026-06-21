"""
Core lead processing engine.
Handles: xlsx, msg_xlsx, msg_body_csv input formats.
Merges by email, builds LeadComments HTML, translates non-Latin text.
"""

import io, re, csv
import pandas as pd

# ── Template columns (output format) ─────────────────────────────────────────
TEMPLATE_COLS = [
    "Referral","Brand","ReceivedDateTime","FirstName","LastName","ContactTitle","Email",
    "Company","Address","County","City","State","ZipCode","Country","LeadSource1",
    "LeadSource2","LeadSource3","LeadComments","PhoneSupplied","PhSuppliedExtension",
    "PhoneResearched","CSRName","PDF","DUNS","WebAddress","SIC","NAICS","noOfEmployees",
    "ParentName","LineOfBusiness","Linkedin_Title","Linkedin_Link","PQ","Latitude",
    "Longitude","DemoLead","about_me","college_1","college_1_degree","college_1_start",
    "college_1_end","college_2","college_2_degree","college_2_start","college_2_end",
    "month_of_joining","about_experience","searched_on_google","linkedin_city",
    "linkedin_state","linkedin_country",
]

# ── Helpers ───────────────────────────────────────────────────────────────────

def is_non_latin(text: str) -> bool:
    return bool(re.search(r"[^\u0000-\u024F\u1E00-\u1EFF]", str(text)))

def safe_zip(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    v = str(val).strip()
    try:
        return str(int(float(v)))
    except Exception:
        return v

def get_val(row, col):
    if not col:
        return ""
    if isinstance(col, list):
        parts = [str(row.get(c, "") or "").strip() for c in col]
        return ", ".join(p for p in parts if p and p.lower() != "nan")
    v = row.get(col, "")
    return "" if (v is None or (isinstance(v, float) and pd.isna(v))) else str(v).strip()

# ── File readers ──────────────────────────────────────────────────────────────

def read_xlsx(file_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes))

def read_msg_xlsx(file_bytes: bytes) -> pd.DataFrame:
    """Extract embedded xlsx/zip from inside a .msg file."""
    pk = file_bytes.find(b"PK\x03\x04")
    if pk == -1:
        raise ValueError("No embedded Excel found in this .msg file.")
    df = pd.read_excel(io.BytesIO(file_bytes[pk:]))
    # If first row is actually the header (Cadenas style), detect & fix
    if df.columns[0].startswith("Unnamed"):
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
    return df

def read_msg_body_csv(file_bytes: bytes) -> pd.DataFrame:
    """Extract CSV data embedded as text in a .msg body (handles multiline quoted fields)."""
    # Decode with latin1 to preserve all byte values as characters
    text = file_bytes.decode("latin1", errors="replace")

    # Find the start of the CSV header line
    header_match = re.search(
        r'(?:^|\n|\x00)((?:"?(?:Email|Lead Source|First.?Name|ID|Customer)"?|ID)[^\n]*,[^\n]+)',
        text, re.I
    )
    if not header_match:
        raise ValueError("No CSV header found in the .msg body.")

    csv_start = header_match.start(1)
    csv_text = text[csv_start:]

    # Find the end: stop at a long null run (binary content)
    null_run = re.search(r'\x00{10,}', csv_text)
    if null_run:
        csv_text = csv_text[:null_run.start()]

    # Remove non-printable chars except newline/tab
    csv_text = re.sub(r'[^\x09\x0a\x0d\x20-\x7e\x80-\xff]', '', csv_text).strip()

    try:
        reader = csv.DictReader(io.StringIO(csv_text))
        rows = list(reader)
    except Exception as e:
        raise ValueError(f"CSV parse error: {e}")

    if not rows:
        raise ValueError("No data rows found in the .msg CSV.")

    df = pd.DataFrame(rows)
    # Clean column names
    df.columns = [re.sub(r'^["\s]+|["\s]+$', '', c) for c in df.columns]
    # Filter rows where Email column has a valid email
    if "Email" in df.columns:
        df = df[df["Email"].str.contains(r"@", na=False)].reset_index(drop=True)
    return df

# ── LeadComments builder ──────────────────────────────────────────────────────

def build_lead_comments(group_rows, config: dict) -> str:
    intro = config.get("lead_intro", "")
    outro = config.get("lead_outro", "")
    fields = config.get("comment_fields", [])
    html = f"{intro}<br><br>" if intro else ""
    for row in group_rows:
        for label, col in fields:
            val = get_val(row, col)
            html += f"<b>{label}: </b>{val}<br>"
        html += "<br>"
    if outro:
        html += outro
    return html.strip()

# ── Passthrough merge (Nason / AMI — already in template format) ──────────────

def merge_passthrough(df: pd.DataFrame, config: dict) -> pd.DataFrame:
    email_col = config["merge_by"]
    rows = []
    for email, group in df.groupby(email_col, sort=False):
        first = group.iloc[0]
        row = {col: str(first.get(col, "") or "") for col in TEMPLATE_COLS if col in df.columns}
        for col in TEMPLATE_COLS:
            if col not in row:
                row[col] = ""
        # Merge LeadComments by appending
        comments = group["LeadComments"].dropna().astype(str).tolist()
        row["LeadComments"] = "\n\n---\n\n".join(c for c in comments if c.strip())
        rows.append(row)
    return pd.DataFrame(rows, columns=TEMPLATE_COLS)

# ── Nexen customer name cleaner ("Alex Best (12282)" → "Alex Best") ──────────

def clean_nexen_name(val: str):
    return re.sub(r"\s*\(\d+\)\s*$", "", str(val or "")).strip()

# ── Main processor ────────────────────────────────────────────────────────────

def process(file_bytes: bytes, config: dict, translated: dict = None) -> pd.DataFrame:
    fmt = config["input_format"]
    src = config.get("source_type", "")

    # --- Read raw data ---
    if fmt == "xlsx":
        df = read_xlsx(file_bytes)
    elif fmt == "msg_xlsx":
        df = read_msg_xlsx(file_bytes)
    elif fmt == "msg_body_csv":
        df = read_msg_body_csv(file_bytes)
    else:
        raise ValueError(f"Unknown input_format: {fmt}")

    # --- Passthrough (already in template format) ---
    if src == "passthrough":
        return merge_passthrough(df, config)

    col_map = config["col_map"]
    email_col = config["merge_by"]
    rows_out = []

    for email, group in df.groupby(email_col, sort=False):
        first = group.iloc[0]
        group_rows = group.to_dict("records")

        def t(field):
            """Get translated value if available, else raw."""
            if translated and field in translated:
                idx = df[df[email_col] == email].index[0]
                pos = list(df.groupby(email_col, sort=False).groups.keys()).index(email)
                return translated[field].get(pos, get_val(first, col_map.get(field, "")))
            raw_col = col_map.get(field, "")
            if isinstance(raw_col, list):
                return get_val(first, raw_col)
            return get_val(first, raw_col)

        row = {col: "" for col in TEMPLATE_COLS}
        row["Email"]         = str(email).strip()
        row["FirstName"]     = clean_nexen_name(t("FirstName")) if src == "nexen" else t("FirstName")
        row["LastName"]      = t("LastName")
        row["ContactTitle"]  = t("ContactTitle")
        row["Company"]       = t("Company")
        row["City"]          = t("City")
        row["State"]         = get_val(first, col_map.get("State", ""))
        row["ZipCode"]       = safe_zip(get_val(first, col_map.get("ZipCode", "")))
        row["Country"]       = t("Country")
        row["PhoneSupplied"] = get_val(first, col_map.get("PhoneSupplied", ""))
        row["LeadSource1"]   = config.get("lead_source_1", "")
        row["LeadSource2"]   = config.get("lead_source_2", "")

        # Address
        addr_cols = col_map.get("Address", [])
        if isinstance(addr_cols, list):
            parts = [str(get_val(first, c)).strip() for c in addr_cols]
            row["Address"] = ", ".join(p for p in parts if p and p.lower() != "nan")
        else:
            row["Address"] = get_val(first, addr_cols)

        row["LeadComments"] = build_lead_comments(group_rows, config)
        rows_out.append(row)

    return pd.DataFrame(rows_out, columns=TEMPLATE_COLS)


def detect_non_latin_fields(file_bytes: bytes, config: dict) -> dict:
    """Return {field_name: {email_index: value}} for fields needing translation."""
    fmt = config["input_format"]
    if config.get("source_type") == "passthrough":
        return {}
    try:
        if fmt == "xlsx":
            df = read_xlsx(file_bytes)
        elif fmt == "msg_xlsx":
            df = read_msg_xlsx(file_bytes)
        elif fmt == "msg_body_csv":
            df = read_msg_body_csv(file_bytes)
        else:
            return {}
    except Exception:
        return {}

    col_map = config["col_map"]
    email_col = config["merge_by"]
    translate_fields = ["FirstName", "LastName", "Company", "City", "Country", "Address"]
    to_translate = {}

    emails = list(df.groupby(email_col, sort=False).groups.keys())
    for i, email in enumerate(emails):
        group = df[df[email_col] == email]
        first = group.iloc[0]
        for field in translate_fields:
            raw_col = col_map.get(field, "")
            if isinstance(raw_col, list):
                val = ", ".join(str(get_val(first, c)) for c in raw_col)
            else:
                val = get_val(first, raw_col)
            if is_non_latin(val):
                to_translate.setdefault(field, {})[i] = val

    return to_translate


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    buf = io.StringIO()
    df.to_csv(buf, index=False)
    return buf.getvalue().encode("utf-8")
