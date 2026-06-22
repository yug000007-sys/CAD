"""
Leads Agent — Multi-Project CAD Lead Processor
Streamlit app: upload raw file → select project → pick comment fields → edit lead sources → download clean CSV
"""

import json
import streamlit as st
import requests
import pandas as pd

from configs import PROJECTS
from processor import (
    process, detect_non_latin_fields, to_csv_bytes,
    TEMPLATE_COLS, is_non_latin, build_lead_comments
)

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Leads Agent",
    page_icon="⚙️",
    layout="centered",
)

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
        "Groq API Key (for translation)",
        type="password",
        help="Required only for non-Latin text (Korean, Arabic, Chinese…). Free key at console.groq.com",
    )

# ── File upload ───────────────────────────────────────────────────────────────
fmt = config["input_format"]
accept_map = {
    "xlsx":         ["xlsx", "xls"],
    "msg_xlsx":     ["msg"],
    "msg_body_csv": ["msg"],
}
accepted = accept_map.get(fmt, ["xlsx", "msg"])

uploaded = st.file_uploader(
    f"Upload raw file for **{project_name}**",
    type=accepted,
    help=f"Expected format: {fmt}",
)

if not uploaded:
    st.stop()

file_bytes = uploaded.read()
st.success(f"✅ **{uploaded.name}** loaded ({len(file_bytes):,} bytes)")

# ── Non-Latin detection & translation ────────────────────────────────────────
with st.spinner("Scanning for non-Latin text..."):
    to_translate = detect_non_latin_fields(file_bytes, config)

translated = {}

if to_translate:
    st.warning(f"⚠️ Non-Latin text detected in {len(to_translate)} field(s).")
    with st.expander("Fields to translate", expanded=True):
        for field, items in to_translate.items():
            for idx, val in items.items():
                st.write(f"**{field}** (row {idx}): `{val}`")

    if not api_key:
        st.error("Enter your Groq API key in the sidebar to translate. Free key at console.groq.com")
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
                    "model": "llama-3.3-70b-versatile",
                    "max_tokens": 1000,
                    "messages": [
                        {"role": "system", "content": (
                            "Translate non-English/non-Latin text to English. "
                            "If already Latin, return as-is. "
                            "Return ONLY a JSON array in the same order, no explanation, no backticks."
                        )},
                        {"role": "user", "content": f"Input: {json.dumps(all_values)}"},
                    ],
                },
                timeout=30,
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

# ── Lead Source editor ────────────────────────────────────────────────────────
st.divider()
st.subheader("🏷️ Lead Sources")
st.caption("Pre-filled from project defaults — edit if needed for this export.")

ls_col1, ls_col2 = st.columns(2)
lead_source_1 = ls_col1.text_input(
    "LeadSource1",
    value=config.get("lead_source_1", ""),
    key="ls1",
)
lead_source_2 = ls_col2.text_input(
    "LeadSource2",
    value=config.get("lead_source_2", ""),
    key="ls2",
)

# ── LeadComments field picker ─────────────────────────────────────────────────
all_comment_fields = config.get("comment_fields", [])
comment_template   = config.get("comment_template", "default")
show_picker = len(all_comment_fields) > 0 and comment_template == "default"

selected_fields = all_comment_fields

if show_picker:
    st.divider()
    st.subheader("📋 LeadComments — choose fields to include")
    st.caption("Check the fields you want included in the LeadComments column for this export.")

    cols = st.columns(2)
    selected_fields = []
    for i, (label, col_key) in enumerate(all_comment_fields):
        checked = cols[i % 2].checkbox(label, value=True, key=f"field_{i}_{label}")
        if checked:
            selected_fields.append((label, col_key))

    if not selected_fields:
        st.warning("⚠️ No fields selected — LeadComments will be empty.")

# ── Process ───────────────────────────────────────────────────────────────────
st.divider()
with st.spinner("Processing leads..."):
    active_config = {
        **config,
        "comment_fields": selected_fields,
        "lead_source_1":  lead_source_1,
        "lead_source_2":  lead_source_2,
    }
    try:
        result_df = process(file_bytes, active_config, translated if translated else None)
    except Exception as e:
        st.error(f"Processing error: {e}")
        st.stop()

# ── Stats ─────────────────────────────────────────────────────────────────────
c1, c2, c3 = st.columns(3)
c1.metric("Unique contacts", len(result_df))
c2.metric("LeadSource1", lead_source_1 or "—")
c3.metric("LeadSource2", lead_source_2 or "—")

# ── Preview ───────────────────────────────────────────────────────────────────
with st.expander("Preview contacts (first 5 rows)"):
    preview_cols = ["Email", "FirstName", "LastName", "Company", "City", "Country", "LeadSource1", "LeadSource2"]
    st.dataframe(result_df[[c for c in preview_cols if c in result_df.columns]].head())

with st.expander("Preview LeadComments (first contact)"):
    if len(result_df):
        st.markdown(result_df["LeadComments"].iloc[0], unsafe_allow_html=True)

# ── Download ──────────────────────────────────────────────────────────────────
fname_base = uploaded.name.rsplit(".", 1)[0]
out_name   = f"{project_name.replace(' ', '_')}_{fname_base}_leads.csv"

st.download_button(
    label="⬇️ Download CSV",
    data=to_csv_bytes(result_df),
    file_name=out_name,
    mime="text/csv",
    use_container_width=True,
)

st.divider()
st.caption("To add a new project, add a config block to `configs.py` — no other code changes needed.")
