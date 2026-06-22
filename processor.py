"""
Core lead processing engine.
Handles: xlsx, csv, msg_xlsx, msg_body_csv input formats.
Auto-detects format from file extension when source_type is 'universal'.
Merges by email, builds per-project LeadComments HTML, translates non-Latin text.
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

def read_csv(file_bytes: bytes) -> pd.DataFrame:
    for enc in ("utf-8", "latin1", "cp1252"):
        try:
            return pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
        except Exception:
            continue
    raise ValueError("Could not read CSV file.")

def read_xlsx(file_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes))

def read_msg_xlsx(file_bytes: bytes) -> pd.DataFrame:
    pk = file_bytes.find(b"PK\x03\x04")
    if pk == -1:
        raise ValueError("No embedded Excel found in this .msg file.")
    df = pd.read_excel(io.BytesIO(file_bytes[pk:]))
    if str(df.columns[0]).startswith("Unnamed"):
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
    return df

def read_msg_body_csv(file_bytes: bytes) -> pd.DataFrame:
    text = file_bytes.decode("latin1", errors="replace")
    header_match = re.search(
        r'(?:^|\n|\x00)((?:"?(?:Email|Lead Source|First.?Name|ID|Customer)"?|ID)[^\n]*,[^\n]+)',
        text, re.I
    )
    if not header_match:
        raise ValueError("No CSV header found in the .msg body.")
    csv_start = header_match.start(1)
    csv_text = text[csv_start:]
    null_run = re.search(r'\x00{10,}', csv_text)
    if null_run:
        csv_text = csv_text[:null_run.start()]
    csv_text = re.sub(r'[^\x09\x0a\x0d\x20-\x7e\x80-\xff]', '', csv_text).strip()
    try:
        reader = csv.DictReader(io.StringIO(csv_text))
        rows = list(reader)
    except Exception as e:
        raise ValueError(f"CSV parse error: {e}")
    if not rows:
        raise ValueError("No data rows found in the .msg CSV.")
    df = pd.DataFrame(rows)
    df.columns = [re.sub(r'^["\s]+|["\s]+$', '', c) for c in df.columns]
    if "Email" in df.columns:
        df = df[df["Email"].str.contains(r"@", na=False)].reset_index(drop=True)
    return df

def auto_map_contact_fields(columns: list) -> dict:
    """
    For Universal mode: fuzzy-match raw column names to standard contact fields.
    Returns a col_map dict like the ones in configs.py.
    """
    cols_lower = {c.lower().strip(): c for c in columns}

    def find(candidates):
        for c in candidates:
            if c in cols_lower:
                return cols_lower[c]
        return ""

    # Address: collect all address-like columns as a list
    addr_cols = []
    for c in columns:
        cl = c.lower().strip()
        if re.match(r"address\s*[123]?$", cl) or cl in ("street", "address1", "address2", "addr1", "addr2", "street address"):
            addr_cols.append(c)

    return {
        "FirstName":     find(["firstname", "first name", "first_name", "fname", "name", "customer name", "contact first name"]),
        "LastName":      find(["lastname", "last name", "last_name", "lname", "surname", "contact last name"]),
        "ContactTitle":  find(["jobtitle", "job title", "job_title", "title", "contacttitle", "contact title", "position", "role"]),
        "Company":       find(["company", "companyname", "company name", "company_name", "organization", "firm", "account", "business"]),
        "Address":       addr_cols if addr_cols else find(["address", "street", "addr"]),
        "City":          find(["city", "town", "municipality"]),
        "State":         find(["state", "state/province", "province", "region", "state (usa)"]),
        "ZipCode":       find(["zip", "zipcode", "zip code", "zip_code", "postal code", "postalcode", "postal_code", "postcode"]),
        "Country":       find(["country", "country name", "countryname", "nation"]),
        "PhoneSupplied": find(["phone", "phonenumber", "phone number", "phone_number", "telephone", "tel", "mobile", "cell"]),
    }



    """Auto-detect format from filename extension."""
    ext = filename.lower().rsplit(".", 1)[-1]
    if ext in ("xlsx", "xls"):
        return read_xlsx(file_bytes)
    elif ext == "csv":
        return read_csv(file_bytes)
    elif ext == "msg":
        # Try embedded xlsx first, fall back to CSV in body
        pk = file_bytes.find(b"PK\x03\x04")
        if pk != -1:
            try:
                return read_msg_xlsx(file_bytes)
            except Exception:
                pass
        return read_msg_body_csv(file_bytes)
    else:
        raise ValueError(f"Unsupported file type: .{ext}")

def auto_read(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Auto-detect format from filename extension."""
    ext = filename.lower().rsplit(".", 1)[-1]
    if ext in ("xlsx", "xls"):
        return read_xlsx(file_bytes)
    elif ext == "csv":
        return read_csv(file_bytes)
    elif ext == "msg":
        pk = file_bytes.find(b"PK\x03\x04")
        if pk != -1:
            try:
                return read_msg_xlsx(file_bytes)
            except Exception:
                pass
        return read_msg_body_csv(file_bytes)
    else:
        raise ValueError(f"Unsupported file type: .{ext}")


def get_columns(file_bytes: bytes, filename: str) -> list:
    """Return list of column names from a file — used by Universal project."""
    df = auto_read(file_bytes, filename)
    return list(df.columns)

# ── LeadComments builders ─────────────────────────────────────────────────────

def build_lead_comments_nexen(group_rows, config: dict) -> str:
    """Nexen format: grouped unique products + content types on single lines."""
    intro  = config.get("lead_intro", "This lead is generated from CAD Download:")
    outro  = config.get("lead_outro", "")
    fields = config.get("comment_fields", [])
    field_map = {label: col for label, col in fields}
    selected_labels = [label for label, _ in fields]

    def unique_vals(col):
        seen, out = set(), []
        for row in group_rows:
            v = get_val(row, col).strip()
            if v and v not in seen:
                seen.add(v)
                out.append(v)
        return out

    html = f"{intro}<br>"
    if "Part Number" in selected_labels:
        vals = unique_vals(field_map.get("Part Number", "Product"))
        if vals:
            html += f"<br><b>Part Number: </b>{', '.join(vals)}"
    if "Content Type" in selected_labels:
        vals = unique_vals(field_map.get("Content Type", "Content Type"))
        if vals:
            html += f"<br><b>Content Type: </b>{', '.join(vals)}"
    if "Date" in selected_labels:
        vals = unique_vals(field_map.get("Date", "Accessed At"))
        if vals:
            html += f"<br><b>Date: </b>{', '.join(vals)}"
    if outro:
        html += f"<br><br>{outro}"
    return html.strip()

def build_lead_comments_default(group_rows, config: dict) -> str:
    """Standard HTML comment block — one block per download row."""
    intro  = config.get("lead_intro", "")
    outro  = config.get("lead_outro", "")
    fields = config.get("comment_fields", [])
    html = f"{intro}<br><br>" if intro else ""
    for row in group_rows:
        for label, col in fields:
            val = get_val(row, col)
            if val:
                html += f"<b>{label}: </b>{val}<br>"
        html += "<br>"
    if outro:
        html += outro
    return html.strip()

def build_lead_comments_nason(group_rows, config: dict) -> str:
    """Nason format: MODEL NUMBER:{model} and cad name: {cad}"""
    intro = config.get("lead_intro", "This is a registered user who has downloaded the cad drawing for")
    outro = config.get("lead_outro", "please contact the customer for service and product opportunities.")
    html  = f"{intro} <br><br>"
    for row in group_rows:
        model = get_val(row, "CAD name") or get_val(row, "Part number") or ""
        cad   = get_val(row, "Standardname") or get_val(row, "CAD format") or ""
        if model:
            html += f"MODEL NUMBER:{model} and cad name: {cad}<br><br>"
    html += outro
    return html.strip()

def build_lead_comments_leak_defense(group_rows, config: dict) -> str:
    """Leak Defense: Notes to Rep IS the body; other fields as labels."""
    fields = config.get("comment_fields", [])
    parts  = []
    for row in group_rows:
        for label, col in fields:
            val = get_val(row, col)
            if val and label == "Notes":
                parts.append(val)
            elif val:
                parts.append(f"<b>{label}: </b>{val}<br>")
    return "<br>".join(parts).strip()

def build_lead_comments(group_rows, config: dict) -> str:
    template = config.get("comment_template", "default")
    src      = config.get("source_type", "")
    if template == "nason" or src == "nason":
        return build_lead_comments_nason(group_rows, config)
    elif template == "leak_defense" or src == "leak_defense":
        return build_lead_comments_leak_defense(group_rows, config)
    elif template == "nexen" or src == "nexen":
        return build_lead_comments_nexen(group_rows, config)
    else:
        return build_lead_comments_default(group_rows, config)

# ── Passthrough merge ─────────────────────────────────────────────────────────

def merge_passthrough(df: pd.DataFrame, config: dict) -> pd.DataFrame:
    email_col = config["merge_by"]
    rows = []
    for email, group in df.groupby(email_col, sort=False):
        first = group.iloc[0]
        row = {}
        for col in TEMPLATE_COLS:
            row[col] = str(first.get(col, "") or "") if col in df.columns else ""
        comments = group["LeadComments"].dropna().astype(str).tolist() if "LeadComments" in group else []
        row["LeadComments"] = "\n\n---\n\n".join(c for c in comments if c.strip())
        rows.append(row)
    return pd.DataFrame(rows, columns=TEMPLATE_COLS)

# ── Nexen name cleaner ────────────────────────────────────────────────────────

def clean_nexen_name(val: str) -> str:
    return re.sub(r"\s*\(\d+\)\s*$", "", str(val or "")).strip()

# ── Smart format resolver ─────────────────────────────────────────────────────

def resolve_df(file_bytes: bytes, config: dict, filename: str = "") -> pd.DataFrame:
    """Read file using config format, with auto fallback for universal projects."""
    fmt = config.get("input_format", "auto")
    src = config.get("source_type", "")

    if fmt == "auto" or src == "universal":
        return auto_read(file_bytes, filename)
    elif fmt == "xlsx":
        return read_xlsx(file_bytes)
    elif fmt == "csv":
        return read_csv(file_bytes)
    elif fmt == "msg_xlsx":
        return read_msg_xlsx(file_bytes)
    elif fmt == "msg_body_csv":
        return read_msg_body_csv(file_bytes)
    else:
        raise ValueError(f"Unknown input_format: {fmt}")

# ── Main processor ────────────────────────────────────────────────────────────

def process(file_bytes: bytes, config: dict, translated: dict = None, filename: str = "") -> pd.DataFrame:
    src = config.get("source_type", "")
    df  = resolve_df(file_bytes, config, filename)

    if src == "passthrough":
        return merge_passthrough(df, config)

    col_map   = config.get("col_map", {})
    email_col = config.get("merge_by", "Email")

    # Universal mode: auto-map contact fields only if no manual map was provided
    if src == "universal" and not col_map:
        col_map = auto_map_contact_fields(list(df.columns))

    rows_out  = []
    email_list = list(df.groupby(email_col, sort=False).groups.keys())

    for pos, email in enumerate(email_list):
        group      = df[df[email_col] == email]
        first      = group.iloc[0]
        group_rows = group.to_dict("records")

        def t(field):
            if translated and field in translated and pos in translated[field]:
                return translated[field][pos]
            raw_col = col_map.get(field, "")
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
        row["LeadSource3"]   = config.get("lead_source_3", "")

        addr_cols = col_map.get("Address", [])
        if isinstance(addr_cols, list):
            parts = [str(get_val(first, c)).strip() for c in addr_cols]
            row["Address"] = ", ".join(p for p in parts if p and p.lower() != "nan")
        else:
            row["Address"] = get_val(first, addr_cols)

        row["LeadComments"] = build_lead_comments(group_rows, config)
        rows_out.append(row)

    return pd.DataFrame(rows_out, columns=TEMPLATE_COLS)


def detect_non_latin_fields(file_bytes: bytes, config: dict, filename: str = "") -> dict:
    if config.get("source_type") == "passthrough":
        return {}
    try:
        df = resolve_df(file_bytes, config, filename)
    except Exception:
        return {}

    col_map   = config.get("col_map", {})
    email_col = config.get("merge_by", "Email")
    translate_fields = ["FirstName", "LastName", "Company", "City", "Country", "Address"]
    to_translate = {}
    emails = list(df.groupby(email_col, sort=False).groups.keys())

    for i, email in enumerate(emails):
        first = df[df[email_col] == email].iloc[0]
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
    """Save CSV with UTF-8 BOM so Excel opens it correctly without encoding corruption."""
    buf = io.BytesIO()
    df.to_csv(buf, index=False, encoding='utf-8-sig')
    return buf.getvalue()
