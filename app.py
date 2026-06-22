"""
Leads Agent — Multi-Project CAD Lead Processor
"""

import json
import streamlit as st
import requests
import pandas as pd

from configs import PROJECTS
from processor import (
    process, detect_non_latin_fields, to_csv_bytes,
    TEMPLATE_COLS, is_non_latin, get_columns, auto_read
)

st.set_page_config(page_title="Leads Agent", page_icon="⚙️", layout="centered")
st.title("⚙️ Leads Agent")
st.caption("Upload a raw lead file → select project → pick comment fields → download clean CSV")

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Project")
    project_name = st.selectbox("Select project", list(PROJECTS.keys()))
    config = PROJECTS[project_name]
    st.info(config["description"])
    st.divider()
    st.header("Settings")
    api_key = st.text_input(
        "Groq API Key (for translation)", type="password",
        help="Required only for non-Latin text. Free key at console.groq.com",
    )

is_universal = config.get("source_type") == "universal"

# ── File upload — all projects accept msg, xlsx, csv ─────────────────────────
uploaded = st.file_uploader(
    f"Upload raw file for **{project_name}**",
    type=["xlsx", "xls", "csv", "msg"],
    help="Accepted: .xlsx, .xls, .csv, .msg",
)

if not uploaded:
    st.stop()

file_bytes = uploaded.read()
filename   = uploaded.name
st.success(f"✅ **{filename}** loaded ({len(file_bytes):,} bytes)")

# ── Non-Latin detection & translation ────────────────────────────────────────
if not is_universal:
    with st.spinner("Scanning for non-Latin text..."):
        to_translate = detect_non_latin_fields(file_bytes, config, filename)

    translated = {}
    if to_translate:
        st.warning(f"⚠️ Non-Latin text detected in {len(to_translate)} field(s).")
        with st.expander("Fields to translate", expanded=True):
            for field, items in to_translate.items():
                for idx, val in items.items():
                    st.write(f"**{field}** (row {idx}): `{val}`")
        if not api_key:
            st.error("Enter your Groq API key in the sidebar. Free key at console.groq.com")
            st.stop()
        with st.spinner("Translating via Groq..."):
            all_values, index_map = [], []
            for field, items in to_translate.items():
                for idx, val in items.items():
                    all_values.append(val)
                    index_map.append((field, idx))
            try:
                resp = requests.post(
                    "https://api.groq.com/openai/v1/chat/completions",
                    headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
                    json={
                        "model": "llama-3.3-70b-versatile", "max_tokens": 1000,
                        "messages": [
                            {"role": "system", "content": (
                                "Translate non-English/non-Latin text to English. "
                                "If already Latin, return as-is. "
                                "Return ONLY a JSON array in the same order, no explanation, no backticks."
                            )},
                            {"role": "user", "content": f"Input: {json.dumps(all_values)}"},
                        ],
                    }, timeout=30,
                )
                trans_list = json.loads(
                    resp.json()["choices"][0]["message"]["content"]
                    .replace("```json", "").replace("```", "").strip()
                )
                for (field, idx), val in zip(index_map, trans_list):
                    translated.setdefault(field, {})[idx] = val
                st.success("✅ Translation complete")
                with st.expander("Translation results"):
                    for (field, idx), val in zip(index_map, trans_list):
                        st.write(f"**{field}**: `{to_translate[field][idx]}` → `{val}`")
            except Exception as e:
                st.error(f"Translation failed: {e}. Proceeding with originals.")
else:
    translated = {}

# ── Lead Sources ──────────────────────────────────────────────────────────────
st.divider()
st.subheader("🏷️ Lead Sources")
st.caption("Pre-filled from project defaults — edit if needed for this export.")

ls_col1, ls_col2, ls_col3 = st.columns(3)
lead_source_1 = ls_col1.text_input("LeadSource1", value=config.get("lead_source_1", ""), key="ls1")
lead_source_2 = ls_col2.text_input("LeadSource2", value=config.get("lead_source_2", ""), key="ls2")
lead_source_3 = ls_col3.text_input("LeadSource3", value=config.get("lead_source_3", ""), key="ls3")

# ── UNIVERSAL: full manual configuration ─────────────────────────────────────
if is_universal:
    st.divider()
    st.subheader("⚙️ Universal Configuration")

    # Detect columns from uploaded file
    try:
        detected_cols = get_columns(file_bytes, filename)
    except Exception as e:
        st.error(f"Could not read file columns: {e}")
        st.stop()

    st.caption(f"Detected {len(detected_cols)} columns from your file.")

    # Contact field manual mapper
    from processor import auto_map_contact_fields
    auto_map = auto_map_contact_fields(detected_cols)

    st.markdown("**📌 Contact field mapping**")
    st.caption("Auto-matched from your columns — override any field using the dropdowns.")

    CONTACT_FIELDS = [
        ("FirstName",    "First Name"),
        ("LastName",     "Last Name"),
        ("ContactTitle", "Job Title"),
        ("Company",      "Company"),
        ("Address",      "Address"),
        ("City",         "City"),
        ("State",        "State"),
        ("ZipCode",      "Zip / Postal Code"),
        ("Country",      "Country"),
        ("PhoneSupplied","Phone"),
    ]

    none_option = "— skip —"
    col_options = [none_option] + detected_cols
    manual_col_map = {}

    grid1, grid2 = st.columns(2)
    for i, (field_key, field_label) in enumerate(CONTACT_FIELDS):
        auto_val = auto_map.get(field_key, "")
        if isinstance(auto_val, list):
            auto_val = auto_val[0] if auto_val else ""
        default_idx = col_options.index(auto_val) if auto_val in col_options else 0
        container = grid1 if i % 2 == 0 else grid2
        chosen = container.selectbox(
            field_label,
            options=col_options,
            index=default_idx,
            key=f"cmap_{field_key}",
        )
        if chosen != none_option:
            if field_key == "Address":
                manual_col_map[field_key] = [chosen]
            else:
                manual_col_map[field_key] = chosen

    # Email column picker
    email_options = [c for c in detected_cols if "email" in c.lower()] + \
                    [c for c in detected_cols if "email" not in c.lower()]
    email_col_pick = st.selectbox("Email column (merge key)", email_options, key="email_col")

    # Intro / Outro
    st.markdown("**LeadComments text**")
    ic1, ic2 = st.columns(2)
    lead_intro = ic1.text_area("Intro (start of comment)", value="", height=80, key="intro")
    lead_outro = ic2.text_area("Outro (end of comment)",   value="", height=80, key="outro")

    # Column picker table with editable labels
    st.markdown("**Select columns for LeadComments & rename labels**")
    st.caption("Check columns to include. Edit the Label column to rename them in the output.")

    selected_fields = []
    header_cols = st.columns([0.5, 3, 3])
    header_cols[0].markdown("**✓**")
    header_cols[1].markdown("**Raw Column**")
    header_cols[2].markdown("**Label in Comment**")

    for i, col_name in enumerate(detected_cols):
        if col_name == email_col_pick:
            continue  # skip the email column
        row_cols = st.columns([0.5, 3, 3])
        checked = row_cols[0].checkbox("", key=f"u_chk_{i}", label_visibility="collapsed")
        row_cols[1].markdown(f"`{col_name}`")
        label = row_cols[2].text_input("", value=col_name, key=f"u_lbl_{i}",
                                        label_visibility="collapsed")
        if checked:
            selected_fields.append((label, col_name))

    # Build active config for universal
    active_config = {
        **config,
        "merge_by":       email_col_pick,
        "col_map":        manual_col_map,
        "comment_fields": selected_fields,
        "comment_template": "default",
        "lead_intro":     lead_intro,
        "lead_outro":     lead_outro,
        "lead_source_1":  lead_source_1,
        "lead_source_2":  lead_source_2,
        "lead_source_3":  lead_source_3,
        "input_format":   "auto",
    }

# ── STANDARD: field picker (checkbox + editable label) ───────────────────────
else:
    all_comment_fields = config.get("comment_fields", [])
    comment_template   = config.get("comment_template", "default")
    show_picker = len(all_comment_fields) > 0 and comment_template not in ("nason", "leak_defense")

    selected_fields = all_comment_fields

    if show_picker:
        st.divider()
        st.subheader("📋 LeadComments — choose fields to include")
        st.caption("Check fields to include. Edit the label to rename them in the output.")

        header_cols = st.columns([0.5, 3, 3])
        header_cols[0].markdown("**✓**")
        header_cols[1].markdown("**Field**")
        header_cols[2].markdown("**Label in Comment**")

        selected_fields = []
        for i, (default_label, col_key) in enumerate(all_comment_fields):
            row_cols = st.columns([0.5, 3, 3])
            checked = row_cols[0].checkbox("", value=True, key=f"s_chk_{i}",
                                            label_visibility="collapsed")
            row_cols[1].markdown(f"`{col_key}`")
            label = row_cols[2].text_input("", value=default_label, key=f"s_lbl_{i}",
                                            label_visibility="collapsed")
            if checked:
                selected_fields.append((label, col_key))

        if not selected_fields:
            st.warning("⚠️ No fields selected — LeadComments will be empty.")

    # Determine input format — auto-detect if msg uploaded for non-msg project
    resolved_fmt = config["input_format"]
    if filename.lower().endswith(".msg") and resolved_fmt in ("xlsx", "csv"):
        resolved_fmt = "auto"
    elif filename.lower().endswith(".csv") and resolved_fmt in ("xlsx", "msg_xlsx", "msg_body_csv"):
        resolved_fmt = "csv"
    elif filename.lower().endswith((".xlsx", ".xls")) and resolved_fmt in ("msg_xlsx", "msg_body_csv", "csv"):
        resolved_fmt = "xlsx"

    active_config = {
        **config,
        "input_format":   resolved_fmt,
        "comment_fields": selected_fields,
        "lead_source_1":  lead_source_1,
        "lead_source_2":  lead_source_2,
        "lead_source_3":  lead_source_3,
    }

# ── Process ───────────────────────────────────────────────────────────────────
st.divider()
with st.spinner("Processing leads..."):
    try:
        result_df = process(file_bytes, active_config, translated if translated else None, filename)
    except Exception as e:
        st.error(f"Processing error: {e}")
        st.stop()

# ── Format ────────────────────────────────────────────────────────────────────
from formatter import format_dataframe

format_steps = []
if result_df["State"].apply(lambda v: bool(str(v).strip()) and len(str(v).strip()) <= 6).any():
    format_steps.append("expanding state abbreviations")
needs_zip_fill = (
    result_df["City"].apply(lambda v: not str(v).strip() or str(v).strip().lower() == "nan") |
    result_df["State"].apply(lambda v: not str(v).strip() or str(v).strip().lower() == "nan")
) & result_df["ZipCode"].apply(lambda v: bool(str(v).strip()) and str(v).strip().lower() != "nan")
if needs_zip_fill.any():
    format_steps.append(f"looking up {needs_zip_fill.sum()} zip codes")

spinner_msg = "Formatting & enriching data" + (f" ({', '.join(format_steps)})" if format_steps else "") + "..."
with st.spinner(spinner_msg):
    try:
        result_df = format_dataframe(result_df, groq_api_key=api_key or "")
    except Exception as e:
        st.warning(f"Formatting partially failed: {e}. Raw data preserved.")

# ── Stats ─────────────────────────────────────────────────────────────────────
c1, c2, c3 = st.columns(3)
c1.metric("Unique contacts", len(result_df))
c2.metric("LeadSource1", lead_source_1 or "—")
c3.metric("LeadSource2", lead_source_2 or "—")

# ── Preview ───────────────────────────────────────────────────────────────────
with st.expander("Preview contacts (first 5 rows)"):
    preview_cols = ["Email","FirstName","LastName","Company","City","Country","LeadSource1","LeadSource2"]
    st.dataframe(result_df[[c for c in preview_cols if c in result_df.columns]].head())

with st.expander("Preview LeadComments (first contact)"):
    if len(result_df):
        st.markdown(result_df["LeadComments"].iloc[0], unsafe_allow_html=True)

# ── Download ──────────────────────────────────────────────────────────────────
fname_base = filename.rsplit(".", 1)[0]
out_name   = f"{project_name.replace(' ', '_')}_{fname_base}_leads.csv"

st.download_button(
    label="⬇️ Download CSV",
    data=to_csv_bytes(result_df),
    file_name=out_name,
    mime="text/csv",
    use_container_width=True,
)

st.divider()
st.caption("To add a new project permanently, add a config block to `configs.py`.")
