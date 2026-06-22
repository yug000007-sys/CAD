"""
Core lead processing engine.
Handles: xlsx, msg_xlsx, msg_body_csv input formats.
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

# ── LeadComments builders ─────────────────────────────────────────────────────

def build_lead_comments_nexen(group_rows, config: dict) -> str:
    """Nexen format: grouped unique products + content types on single lines."""
    intro = config.get("lead_intro", "This lead is generated from CAD Download:")
    outro = config.get("lead_outro", "")

    # Collect selected comment fields from config
    fields = config.get("comment_fields", [])
    field_map = {label: col for label, col in fields}

    # Get selected field labels (respects field picker)
    product_col      = field_map.get("Part Number",   "Product")
    content_type_col = field_map.get("Content Type",  "Content Type")
    date_col         = field_map.get("Date",           "Accessed At")

    # Deduplicated values preserving order
    def unique_vals(col):
        seen, out = set(), []
        for row in group_rows:
            v = get_val(row, col).strip()
            if v and v not in seen:
                seen.add(v)
                out.append(v)
        return out

    html = f"{intro}<br>"

    # Only include lines for fields that are selected
    selected_labels = [label for label, _ in fields]

    if "Part Number" in selected_labels:
        products = unique_vals(product_col)
        if products:
            html += f"<br><b>Part Number: </b>{', '.join(products)}"

    if "Content Type" in selected_labels:
        ctypes = unique_vals(content_type_col)
        if ctypes:
            html += f"<br><b>Content Type: </b>{', '.join(ctypes)}"

    if "Date" in selected_labels:
        dates = unique_vals(date_col)
        if dates:
            html += f"<br><b>Date: </b>{', '.join(dates)}"

    if outro:
        html += f"<br><br>{outro}"

    return html.strip()


def build_lead_comments_default(group_rows, config: dict) -> str:
    """Standard HTML comment block used by most projects."""
    intro = config.get("lead_intro", "")
    outro = config.get("lead_outro", "")
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
    html = f"{intro} <br><br>"
    for row in group_rows:
        model = get_val(row, "CAD name") or get_val(row, "Part number") or ""
        cad   = get_val(row, "Standardname") or get_val(row, "CAD format") or ""
        if model:
            html += f"MODEL NUMBER:{model} and cad name: {cad}<br><br>"
    html += outro
    return html.strip()

def build_lead_comments_leak_defense(group_rows, config: dict) -> str:
    """Leak Defense: Notes to Rep field IS the comment body, plus lead source info."""
    fields = config.get("comment_fields", [])
    parts = []
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
        # Merge LeadComments
        comments = group["LeadComments"].dropna().astype(str).tolist() if "LeadComments" in group else []
        row["LeadComments"] = "\n\n---\n\n".join(c for c in comments if c.strip())
        rows.append(row)
    return pd.DataFrame(rows, columns=TEMPLATE_COLS)

# ── Nexen name cleaner ────────────────────────────────────────────────────────

def clean_nexen_name(val: str) -> str:
    return re.sub(r"\s*\(\d+\)\s*$", "", str(val or "")).strip()

# ── Main processor ────────────────────────────────────────────────────────────

def process(file_bytes: bytes, config: dict, translated: dict = None) -> pd.DataFrame:
    fmt = config["input_format"]
    src = config.get("source_type", "")

    if fmt == "xlsx":
        df = read_xlsx(file_bytes)
    elif fmt == "csv":
        df = read_csv(file_bytes)
    elif fmt == "msg_xlsx":
        df = read_msg_xlsx(file_bytes)
    elif fmt == "msg_body_csv":
        df = read_msg_body_csv(file_bytes)
    else:
        raise ValueError(f"Unknown input_format: {fmt}")

    if src == "passthrough":
        return merge_passthrough(df, config)

    col_map = config["col_map"]
    email_col = config["merge_by"]
    rows_out = []
    email_list = list(df.groupby(email_col, sort=False).groups.keys())

    for pos, email in enumerate(email_list):
        group = df[df[email_col] == email]
        first = group.iloc[0]
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
    if config.get("source_type") == "passthrough":
        return {}
    fmt = config["input_format"]
    try:
        if fmt == "xlsx":
            df = read_xlsx(file_bytes)
        elif fmt == "csv":
            df = read_csv(file_bytes)
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
    buf = io.StringIO()
    df.to_csv(buf, index=False)
    return buf.getvalue().encode("utf-8")
