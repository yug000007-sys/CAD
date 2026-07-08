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


def read_msg_csv_attachment(file_bytes: bytes) -> pd.DataFrame:
    """Extract a CSV attachment from a .msg file."""
    try:
        import extract_msg
        msg = extract_msg.Message(io.BytesIO(file_bytes))
        for att in msg.attachments:
            name = (att.longFilename or att.shortFilename or "").lower()
            if name.endswith(".csv"):
                return pd.read_csv(io.BytesIO(att.data))
    except Exception:
        pass
    raise ValueError("No CSV attachment found in this .msg file.")

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

def parse_combined_address(address: str) -> dict:
    """
    Parse a combined address string like:
    '1450 Northeast 138th Avenue, Vancouver, WA 98684, United States'
    into: street, city, state, zip, country
    """
    if not address or str(address).strip().lower() in ("nan", "none", ""):
        return {"street": "", "city": "", "state": "", "zip": "", "country": ""}

    parts = [p.strip() for p in str(address).split(",")]
    result = {"street": "", "city": "", "state": "", "zip": "", "country": ""}

    if len(parts) == 1:
        result["street"] = parts[0]
        return result

    # Last part is usually country
    result["country"] = parts[-1].strip()

    # Second to last: "STATE ZIP" or just city
    if len(parts) >= 3:
        state_zip = parts[-2].strip()
        # Try to match "ST 12345" or "ST 12345-6789" or "QC H2Y 1S1" (Canada) or "CDMX 03100"
        m = re.match(r'^([A-Za-z]{2,4})\s+([A-Z0-9]{3,10}(?:[\s-]\d{4})?)$', state_zip, re.I)
        if m:
            result["state"] = m.group(1).upper()
            result["zip"]   = m.group(2).upper()
        else:
            # maybe just a state or just a zip
            if re.match(r'^[A-Za-z]{2,4}$', state_zip):
                result["state"] = state_zip.upper()
            elif re.match(r'^[\dA-Z]{4,10}$', state_zip, re.I):
                result["zip"] = state_zip.upper()
            else:
                result["city"] = state_zip

    if len(parts) >= 4:
        result["city"]   = parts[-3].strip()
        result["street"] = ", ".join(parts[:-3]).strip()
    elif len(parts) == 3:
        result["city"]   = parts[0].strip()
        result["street"] = ""

    return result


def build_lead_comments_itt_batch(group_rows, config: dict) -> str:
    """
    ITT_Batch format:
    This Informational lead was generated from {event_name}.
    Please contact the customer for product and service opportunities.
    Other Contacts in this organization:-
    Full Name: Name1, Name2, Name3
    E-Mails: email1, email2, email3
    Title: Title1, Title2, Title3
    """
    intro  = config.get("lead_intro", "")
    outro  = config.get("lead_outro", "Please contact the customer for product and service opportunities.")
    col_map = config.get("col_map", {})

    name_col  = col_map.get("_fullname", "Full Name")
    email_col = col_map.get("_email",    "Email")
    title_col = col_map.get("_title",    "Title")

    names  = []
    emails = []
    titles = []

    for row in group_rows:
        n = get_val(row, name_col)
        e = get_val(row, email_col)
        t = get_val(row, title_col)
        if n: names.append(n)
        if e: emails.append(e)
        if t: titles.append(t)

    html  = f"This Informational lead was generated from {intro}.<br>" if intro else ""
    html += f"{outro}<br>"
    html += "Other Contacts in this organization:-<br>"
    if names:
        html += f"<b>Full Name: </b>{', '.join(names)}<br>"
    if emails:
        html += f"<b>E-Mails: </b>{', '.join(emails)}<br>"
    if titles:
        html += f"<b>Title: </b>{', '.join(titles)}<br>"

    return html.strip()



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

def format_watts_comment(text: str) -> str:
    """Convert Notes to Rep text: newlines → <br>, URLs → hyperlinks."""
    if not text or str(text).strip().lower() in ("nan", "none", ""):
        return ""
    text = str(text).replace("\\\\n", "\n").replace("\\n", "\n")
    # Convert URLs to hyperlinks
    text = re.sub(r"(https?://[^\s]+)", r'<a href="\1">\1</a>', text)
    # Convert newlines to <br>
    text = text.replace("\n", "<br>")
    return text


def process_watts_batch(df: pd.DataFrame, config: dict) -> pd.DataFrame:
    """
    Watts Batch: one output row per input row (no deduplication).
    Maps columns directly, formats Notes to Rep as HTML comment.
    """
    col_map = config.get("col_map", {})

    def fc(key):
        """Find column in df by mapped name, case-insensitive."""
        target = col_map.get(key, key)
        cols_lower = {c.lower().strip(): c for c in df.columns}
        return cols_lower.get(target.lower().strip(), target)

    rows_out = []
    for _, row in df.iterrows():
        out = {col: "" for col in TEMPLATE_COLS}
        out["FirstName"]    = str(row.get(fc("FirstName"),   "") or "").strip()
        out["LastName"]     = str(row.get(fc("LastName"),    "") or "").strip()
        out["Email"]        = str(row.get(fc("Email"),       "") or "").strip().lower()
        out["Company"]      = str(row.get(fc("Company"),     "") or "").strip()
        out["City"]         = str(row.get(fc("City"),        "") or "").strip()
        out["State"]        = str(row.get(fc("State"),       "") or "").strip()
        out["Country"]      = str(row.get(fc("Country"),     "") or "").strip()
        out["ZipCode"]      = str(row.get(fc("ZipCode"),     "") or "").strip()
        out["LeadSource1"]  = str(row.get(fc("LeadSource1"), "") or "").strip()
        out["LeadSource2"]  = str(row.get(fc("LeadSource2"), "") or "").strip()
        out["LeadSource3"]  = str(row.get(fc("LeadSource3"), "") or "").strip()
        out["Brand"]        = str(row.get(fc("Brand"),       "") or "").strip()
        out["ContactTitle"] = str(row.get(fc("ContactTitle"),"") or "").strip()
        out["LeadComments"] = format_watts_comment(row.get(fc("LeadComments"), ""))
        rows_out.append(out)

    return pd.DataFrame(rows_out, columns=TEMPLATE_COLS)


def build_lead_comments_default(group_rows, config: dict) -> str:
    """
    Default comment builder — formats selected fields as:
    <b>Label: </b>Value<br>
    """
    intro          = config.get("lead_intro", "")
    outro          = config.get("lead_outro", "")
    selected_fields = config.get("selected_fields", config.get("comment_fields", []))

    lines = []
    if intro:
        lines.append(f"{intro}<br>")

    for row in group_rows:
        for (label, col_key) in selected_fields:
            val = get_val(row, col_key)
            if val:
                lines.append(f"<b>{label}: </b>{val}<br>")
        break  # default: use first row only (standard projects are not grouped)

    if outro:
        lines.append(f"{outro}<br>")

    return "".join(lines).strip()


def build_lead_comments(group_rows, config: dict) -> str:
    template = config.get("comment_template", "default")
    src      = config.get("source_type", "")
    if template == "nason" or src == "nason":
        return build_lead_comments_nason(group_rows, config)
    elif template == "leak_defense" or src == "leak_defense":
        return build_lead_comments_leak_defense(group_rows, config)
    elif template == "nexen" or src == "nexen":
        return build_lead_comments_nexen(group_rows, config)
    elif template == "itt_batch" or src == "itt_batch":
        return build_lead_comments_itt_batch(group_rows, config)
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

    # Watts Batch: msg with CSV attachment, or plain CSV/xlsx
    if src == "watts_batch":
        if filename.lower().endswith(".msg"):
            return read_msg_csv_attachment(file_bytes)
        elif filename.lower().endswith(".csv"):
            return read_csv(file_bytes)
        else:
            return read_xlsx(file_bytes)

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

def process_itt_batch(df: pd.DataFrame, config: dict) -> pd.DataFrame:
    """
    ITT_Batch: groups rows by Company.
    - Parses combined Address field into street/city/state/zip/country
    - Lists all contacts (name, email, title) in LeadComments
    - Contact fields (FirstName, LastName, Email, Phone) left blank
    """
    col_map  = config.get("col_map", {})
    rows_out = []

    # Auto-detect column names if raw file uses different casing
    cols_lower = {c.lower().strip(): c for c in df.columns}
    def find_col(candidates):
        for c in candidates:
            if c.lower() in cols_lower:
                return cols_lower[c.lower()]
        return ""

    company_col = find_col(["company", "company name", "organization"])
    address_col = find_col(["address", "full address", "street address", "addr"])
    name_col    = find_col(["full name", "fullname", "name", "contact name"])
    email_col   = find_col(["email", "e-mail", "email address"])
    title_col   = find_col(["title", "job title", "contact title", "position", "role"])

    # Update config col_map with detected columns
    active_config = {
        **config,
        "col_map": {
            **col_map,
            "Company":   company_col,
            "Address":   address_col,
            "_fullname": name_col,
            "_email":    email_col,
            "_title":    title_col,
        }
    }

    if not company_col or company_col not in df.columns:
        raise ValueError("Could not find 'Company' column in the file.")

    for company, group in df.groupby(company_col, sort=False):
        first = group.iloc[0]
        group_rows = group.to_dict("records")

        # Parse address from first row
        raw_addr = get_val(first, address_col)
        addr = parse_combined_address(raw_addr)

        row = {col: "" for col in TEMPLATE_COLS}
        row["Company"]  = str(company).strip()
        row["Address"]  = addr["street"]
        row["City"]     = addr["city"]
        row["State"]    = addr["state"]
        row["ZipCode"]  = addr["zip"]
        row["Country"]  = addr["country"]
        row["LeadSource1"] = config.get("lead_source_1", "")
        row["LeadSource2"] = config.get("lead_source_2", "")
        row["LeadSource3"] = config.get("lead_source_3", "")
        row["LeadComments"] = build_lead_comments_itt_batch(group_rows, active_config)
        rows_out.append(row)

    return pd.DataFrame(rows_out, columns=TEMPLATE_COLS)


def process(file_bytes: bytes, config: dict, translated: dict = None, filename: str = "") -> pd.DataFrame:
    src = config.get("source_type", "")
    df  = resolve_df(file_bytes, config, filename)

    if src == "passthrough":
        return merge_passthrough(df, config)

    # ITT_Batch: merge by Company, parse combined address, list all contacts in comments
    if src == "itt_batch":
        return process_itt_batch(df, config)

    # Watts Batch: one row per input row, format Notes to Rep as HTML
    if src == "watts_batch":
        return process_watts_batch(df, config)

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
