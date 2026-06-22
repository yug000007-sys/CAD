"""
formatter.py — Data formatting and enrichment for all lead fields.

Rules:
- FirstName, LastName, ContactTitle, Company, Address, City, Country → Proper Case
- Email → lowercase
- Phone → country-aware formatting
  - USA/Canada: xxx-xxx-xxxx (local format, no country code)
  - International: dialcode-number
- State:
  - USA/Canada → abbreviation (from reference file)
  - International → full name expanded via Groq
- Zip → City/State auto-fill when blank:
  - USA → local Excel reference file
  - Canada → Zippopotam.us API
  - International → Groq
"""

import re
import json
import requests
import pandas as pd

# ── Load USA ZIP reference data ───────────────────────────────────────────────
_USA_ZIP_DF = None

def _load_usa_zip():
    global _USA_ZIP_DF
    if _USA_ZIP_DF is None:
        try:
            df = pd.read_excel(
                "USA_Canada_ZIP_Postal_City_State.xlsx",
                sheet_name="USA ZIP City State",
                dtype={"ZIP Code": str},
            )
            df["ZIP Code"] = df["ZIP Code"].str.strip().str.zfill(5)
            _USA_ZIP_DF = df.set_index("ZIP Code")
        except Exception:
            _USA_ZIP_DF = pd.DataFrame()
    return _USA_ZIP_DF

# ── Country helpers ───────────────────────────────────────────────────────────

USA_NAMES  = {"usa", "united states", "united states of america", "us", "u.s.", "u.s.a."}
CA_NAMES   = {"canada", "ca"}

# Country dial codes
DIAL_CODES = {
    "afghanistan": "93", "albania": "355", "algeria": "213", "andorra": "376",
    "angola": "244", "argentina": "54", "armenia": "374", "australia": "61",
    "austria": "43", "azerbaijan": "994", "bahrain": "973", "bangladesh": "880",
    "belarus": "375", "belgium": "32", "belize": "501", "benin": "229",
    "bolivia": "591", "bosnia": "387", "botswana": "267", "brazil": "55",
    "brunei": "673", "bulgaria": "359", "burkina faso": "226", "burundi": "257",
    "cambodia": "855", "cameroon": "237", "chile": "56", "china": "86",
    "colombia": "57", "costa rica": "506", "croatia": "385", "cuba": "53",
    "cyprus": "357", "czech republic": "420", "denmark": "45", "ecuador": "593",
    "egypt": "20", "el salvador": "503", "estonia": "372", "ethiopia": "251",
    "finland": "358", "france": "33", "georgia": "995", "germany": "49",
    "ghana": "233", "greece": "30", "guatemala": "502", "honduras": "504",
    "hungary": "36", "iceland": "354", "india": "91", "indonesia": "62",
    "iran": "98", "iraq": "964", "ireland": "353", "israel": "972",
    "italy": "39", "jamaica": "1", "japan": "81", "jordan": "962",
    "kazakhstan": "7", "kenya": "254", "kuwait": "965", "kyrgyzstan": "996",
    "latvia": "371", "lebanon": "961", "libya": "218", "liechtenstein": "423",
    "lithuania": "370", "luxembourg": "352", "malaysia": "60", "maldives": "960",
    "mali": "223", "malta": "356", "mauritius": "230", "mexico": "52",
    "moldova": "373", "monaco": "377", "mongolia": "976", "montenegro": "382",
    "morocco": "212", "mozambique": "258", "myanmar": "95", "namibia": "264",
    "nepal": "977", "netherlands": "31", "new zealand": "64", "nicaragua": "505",
    "niger": "227", "nigeria": "234", "north korea": "850", "norway": "47",
    "oman": "968", "pakistan": "92", "panama": "507", "paraguay": "595",
    "peru": "51", "philippines": "63", "poland": "48", "portugal": "351",
    "qatar": "974", "romania": "40", "russia": "7", "rwanda": "250",
    "saudi arabia": "966", "senegal": "221", "serbia": "381", "singapore": "65",
    "slovakia": "421", "slovenia": "386", "somalia": "252", "south africa": "27",
    "south korea": "82", "spain": "34", "sri lanka": "94", "sudan": "249",
    "sweden": "46", "switzerland": "41", "syria": "963", "taiwan": "886",
    "tajikistan": "992", "tanzania": "255", "thailand": "66", "togo": "228",
    "trinidad and tobago": "1", "tunisia": "216", "turkey": "90",
    "turkmenistan": "993", "uganda": "256", "ukraine": "380",
    "united arab emirates": "971", "uae": "971", "united kingdom": "44",
    "uk": "44", "uruguay": "598", "uzbekistan": "998", "venezuela": "58",
    "vietnam": "84", "yemen": "967", "zambia": "260", "zimbabwe": "263",
}

def is_usa(country: str) -> bool:
    return country.strip().lower() in USA_NAMES

def is_canada(country: str) -> bool:
    return country.strip().lower() in CA_NAMES

def is_usa_or_canada(country: str) -> bool:
    return is_usa(country) or is_canada(country)

def get_dial_code(country: str) -> str:
    return DIAL_CODES.get(country.strip().lower(), "")

# ── Proper case ───────────────────────────────────────────────────────────────

# Words that should stay lowercase in proper case
_LOWER_WORDS = {"a","an","the","and","but","or","for","nor","on","at","to",
                "by","in","of","up","as","is","it","its","de","del","la","le",
                "les","las","los","von","van","der","den","di","da","e"}

_KEEP_UPPER = {"LLC", "LLP", "PLC", "USA", "UK", "UAE",
               "SRL", "SA", "AG", "BV", "NV", "AB", "AS", "OY", "KG", "GMBH",
               "PTE", "SDN", "BHD", "PTY", "LP"}

# These get title-cased with period preserved
_TITLE_ABBR = {"inc": "Inc.", "corp": "Corp.", "ltd": "Ltd.", "co": "Co.",
               "inc.": "Inc.", "corp.": "Corp.", "ltd.": "Ltd."}

def proper_case(text: str) -> str:
    if not text or str(text).strip().lower() in ("nan", "none", ""):
        return ""
    words = str(text).strip().split()
    result = []
    for i, w in enumerate(words):
        w_clean  = w.rstrip(".,;:")
        w_upper  = w_clean.upper()
        w_lower  = w_clean.lower()
        if w_upper in _KEEP_UPPER:
            result.append(w_upper)
        elif w_lower in _TITLE_ABBR:
            result.append(_TITLE_ABBR[w_lower])
        elif i == 0 or w_lower not in _LOWER_WORDS:
            result.append(w.capitalize())
        else:
            result.append(w.lower())
    return " ".join(result)

# ── Phone formatting ──────────────────────────────────────────────────────────

def _strip_phone(phone: str) -> str:
    """Remove all non-digit characters."""
    return re.sub(r"\D", "", str(phone))

def format_phone(phone: str, country: str) -> str:
    if not phone or str(phone).strip() in ("nan", "none", ""):
        return ""
    digits = _strip_phone(phone)
    if not digits or len(digits) < 7:
        return str(phone).strip()

    country_clean = (country or "").strip().lower()

    if is_usa(country_clean) or country_clean in ("", "nan"):
        # USA: strip leading 1 if present, format as xxx-xxx-xxxx
        if digits.startswith("1") and len(digits) == 11:
            digits = digits[1:]
        if len(digits) == 10:
            return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
        return digits

    elif is_canada(country_clean):
        # Canada: same local format xxx-xxx-xxxx
        if digits.startswith("1") and len(digits) == 11:
            digits = digits[1:]
        if len(digits) == 10:
            return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
        return digits

    else:
        # International: dialcode-localdigits
        dial = get_dial_code(country_clean)
        if not dial:
            return digits  # unknown country, return clean digits

        # Strip country code prefix if already present
        if digits.startswith(dial) and len(digits) > len(dial) + 6:
            local = digits[len(dial):]
        else:
            local = digits

        return f"{dial}-{local}"

# ── State formatting ──────────────────────────────────────────────────────────

# USA state full → abbreviation
USA_STATE_MAP = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR",
    "california": "CA", "colorado": "CO", "connecticut": "CT", "delaware": "DE",
    "florida": "FL", "georgia": "GA", "hawaii": "HI", "idaho": "ID",
    "illinois": "IL", "indiana": "IN", "iowa": "IA", "kansas": "KS",
    "kentucky": "KY", "louisiana": "LA", "maine": "ME", "maryland": "MD",
    "massachusetts": "MA", "michigan": "MI", "minnesota": "MN", "mississippi": "MS",
    "missouri": "MO", "montana": "MT", "nebraska": "NE", "nevada": "NV",
    "new hampshire": "NH", "new jersey": "NJ", "new mexico": "NM",
    "new york": "NY", "north carolina": "NC", "north dakota": "ND",
    "ohio": "OH", "oklahoma": "OK", "oregon": "OR", "pennsylvania": "PA",
    "rhode island": "RI", "south carolina": "SC", "south dakota": "SD",
    "tennessee": "TN", "texas": "TX", "utah": "UT", "vermont": "VT",
    "virginia": "VA", "washington": "WA", "west virginia": "WV",
    "wisconsin": "WI", "wyoming": "WY", "district of columbia": "DC",
    "puerto rico": "PR", "guam": "GU", "virgin islands": "VI",
}
USA_ABBR_SET = set(USA_STATE_MAP.values())

# Canada province full → abbreviation
CA_PROVINCE_MAP = {
    "alberta": "AB", "british columbia": "BC", "manitoba": "MB",
    "new brunswick": "NB", "newfoundland and labrador": "NL",
    "newfoundland": "NL", "labrador": "NL", "northwest territories": "NT",
    "nova scotia": "NS", "nunavut": "NU", "ontario": "ON",
    "prince edward island": "PE", "quebec": "QC", "québec": "QC",
    "saskatchewan": "SK", "yukon": "YT",
}
CA_ABBR_SET = set(CA_PROVINCE_MAP.values())

def format_state(state: str, country: str) -> str:
    if not state or str(state).strip().lower() in ("nan", "none", ""):
        return ""
    s = str(state).strip()

    if is_usa(country):
        sl = s.lower()
        if sl in USA_STATE_MAP:
            return USA_STATE_MAP[sl]
        if s.upper() in USA_ABBR_SET:
            return s.upper()
        # Already an abbreviation we don't recognize — return uppercase
        return s.upper() if len(s) <= 3 else s

    elif is_canada(country):
        sl = s.lower()
        if sl in CA_PROVINCE_MAP:
            return CA_PROVINCE_MAP[sl]
        if s.upper() in CA_ABBR_SET:
            return s.upper()
        return s.upper() if len(s) <= 3 else s

    else:
        # International — return as-is for now; Groq expansion done in batch
        return s

# ── Zip lookup — USA (local file) ────────────────────────────────────────────

def lookup_usa_zip(zipcode: str) -> dict:
    """Return {'city': ..., 'state': ...} for a US ZIP, or empty dict."""
    df = _load_usa_zip()
    if df.empty:
        return {}
    z = str(zipcode).strip().zfill(5)
    if z in df.index:
        row = df.loc[z]
        if isinstance(row, pd.DataFrame):
            row = row.iloc[0]
        return {
            "city":  proper_case(str(row.get("City", ""))),
            "state": str(row.get("State Abbreviation", "")),
        }
    return {}

# ── Zip lookup — Canada (Zippopotam.us) ──────────────────────────────────────

def lookup_canada_zip(postal: str) -> dict:
    """Return {'city': ..., 'state': ...} via Zippopotam.us for Canadian postal codes."""
    if not postal:
        return {}
    prefix = re.sub(r"\s", "", str(postal).upper())[:3]
    try:
        resp = requests.get(
            f"https://api.zippopotam.us/ca/{prefix}",
            timeout=5,
        )
        if resp.status_code == 200:
            data = resp.json()
            places = data.get("places", [])
            if places:
                city  = proper_case(places[0].get("place name", ""))
                state = places[0].get("state abbreviation", "")
                return {"city": city, "state": state}
    except Exception:
        pass
    return {}

# ── Groq helpers ──────────────────────────────────────────────────────────────

def _groq_call(prompt: str, api_key: str) -> str:
    resp = requests.post(
        "https://api.groq.com/openai/v1/chat/completions",
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json={
            "model": "llama-3.3-70b-versatile",
            "max_tokens": 1000,
            "messages": [
                {"role": "system", "content": (
                    "You are a data formatting assistant. "
                    "Return ONLY a JSON object as instructed. No explanation, no markdown."
                )},
                {"role": "user", "content": prompt},
            ],
        },
        timeout=30,
    )
    return resp.json()["choices"][0]["message"]["content"].replace("```json","").replace("```","").strip()

def expand_states_groq(state_country_pairs: list, api_key: str) -> list:
    """
    Expand abbreviated international states to full names via Groq.
    Input: [{"state": "QUE", "country": "Mexico"}, ...]
    Output: ["Querétaro", ...]
    """
    if not state_country_pairs or not api_key:
        return [p["state"] for p in state_country_pairs]
    prompt = (
        "Expand these abbreviated state/province/region names to their full official names. "
        "If already full, return as-is. Use proper accents (e.g. Querétaro not Queretaro).\n"
        f"Return ONLY a JSON array of strings in the same order.\n\n"
        f"Input: {json.dumps(state_country_pairs)}"
    )
    try:
        result = json.loads(_groq_call(prompt, api_key))
        return result if isinstance(result, list) else [p["state"] for p in state_country_pairs]
    except Exception:
        return [p["state"] for p in state_country_pairs]

def lookup_intl_zip_groq(zip_country_pairs: list, api_key: str) -> list:
    """
    Lookup city/state for international zip codes via Groq.
    Input: [{"zip": "06600", "country": "Mexico"}, ...]
    Output: [{"city": "Mexico City", "state": "CDMX"}, ...]
    """
    if not zip_country_pairs or not api_key:
        return [{"city": "", "state": ""} for _ in zip_country_pairs]
    prompt = (
        "For each zip/postal code and country, provide the city and state/province/region. "
        "Use proper case with correct accents. If unknown, return empty strings.\n"
        f"Return ONLY a JSON array of objects with 'city' and 'state' keys, same order.\n\n"
        f"Input: {json.dumps(zip_country_pairs)}"
    )
    try:
        result = json.loads(_groq_call(prompt, api_key))
        if isinstance(result, list) and len(result) == len(zip_country_pairs):
            return result
    except Exception:
        pass
    return [{"city": "", "state": ""} for _ in zip_country_pairs]

# ── Main batch formatter ──────────────────────────────────────────────────────

def format_dataframe(df: pd.DataFrame, groq_api_key: str = "") -> pd.DataFrame:
    """
    Apply all formatting rules to a processed output DataFrame.
    Returns a new DataFrame with formatted values.
    """
    df = df.copy()

    def safe(val):
        v = str(val).strip()
        return "" if v.lower() in ("nan", "none", "") else v

    # ── 1. Proper case fields ────────────────────────────────────────────────
    for col in ["FirstName", "LastName", "ContactTitle", "Company", "Address", "City", "Country"]:
        if col in df.columns:
            df[col] = df[col].apply(lambda v: proper_case(safe(v)))

    # ── 2. Email lowercase ───────────────────────────────────────────────────
    if "Email" in df.columns:
        df["Email"] = df["Email"].apply(lambda v: safe(v).lower())

    # ── 3. Phone formatting ──────────────────────────────────────────────────
    if "PhoneSupplied" in df.columns:
        df["PhoneSupplied"] = df.apply(
            lambda row: format_phone(safe(row["PhoneSupplied"]), safe(row.get("Country", ""))),
            axis=1,
        )

    # ── 4. State formatting ──────────────────────────────────────────────────
    if "State" in df.columns:
        df["State"] = df.apply(
            lambda row: format_state(safe(row["State"]), safe(row.get("Country", ""))),
            axis=1,
        )

    # ── 5. Batch expand international state abbreviations via Groq ───────────
    if "State" in df.columns and groq_api_key:
        intl_mask = df.apply(
            lambda row: bool(safe(row["State"])) and
                        not is_usa_or_canada(safe(row.get("Country", ""))) and
                        len(safe(row["State"])) <= 6,
            axis=1,
        )
        if intl_mask.any():
            pairs = df[intl_mask][["State", "Country"]].rename(
                columns={"State": "state", "Country": "country"}
            ).to_dict("records")
            expanded = expand_states_groq(pairs, groq_api_key)
            df.loc[intl_mask, "State"] = expanded

    # ── 6. Zip → City/State auto-fill ───────────────────────────────────────
    needs_city  = df["City"].apply(lambda v: not safe(v))  if "City"  in df.columns else pd.Series([False]*len(df))
    needs_state = df["State"].apply(lambda v: not safe(v)) if "State" in df.columns else pd.Series([False]*len(df))
    needs_fill  = (needs_city | needs_state) & df.get("ZipCode", pd.Series([""] * len(df))).apply(lambda v: bool(safe(str(v))))

    if needs_fill.any():
        usa_mask    = needs_fill & df["Country"].apply(lambda v: is_usa(safe(v)))
        canada_mask = needs_fill & df["Country"].apply(lambda v: is_canada(safe(v)))
        intl_mask2  = needs_fill & ~usa_mask & ~canada_mask

        # USA lookup (local file)
        for idx in df[usa_mask].index:
            result = lookup_usa_zip(safe(str(df.at[idx, "ZipCode"])))
            if result:
                if not safe(df.at[idx, "City"]):
                    df.at[idx, "City"]  = result.get("city", "")
                if not safe(df.at[idx, "State"]):
                    df.at[idx, "State"] = result.get("state", "")

        # Canada lookup (Zippopotam.us)
        for idx in df[canada_mask].index:
            result = lookup_canada_zip(safe(str(df.at[idx, "ZipCode"])))
            if result:
                if not safe(df.at[idx, "City"]):
                    df.at[idx, "City"]  = result.get("city", "")
                if not safe(df.at[idx, "State"]):
                    df.at[idx, "State"] = result.get("state", "")

        # International lookup (Groq)
        if intl_mask2.any() and groq_api_key:
            intl_rows = df[intl_mask2]
            pairs = [
                {"zip": safe(str(row["ZipCode"])), "country": safe(row.get("Country", ""))}
                for _, row in intl_rows.iterrows()
            ]
            results = lookup_intl_zip_groq(pairs, groq_api_key)
            for (idx, _), res in zip(intl_rows.iterrows(), results):
                if not safe(df.at[idx, "City"]):
                    df.at[idx, "City"]  = proper_case(res.get("city", ""))
                if not safe(df.at[idx, "State"]):
                    df.at[idx, "State"] = res.get("state", "")

    return df
