"""
Leads Agent — Multi-Project CAD Lead Processor
Streamlit app: upload raw file → select project → download clean CSV
"""

import json
import streamlit as st
import requests
import pandas as pd

# Flat imports — works on Streamlit Cloud
from configs import PROJECTS
from processor import (
    process, detect_non_latin_fields, to_csv_bytes, TEMPLATE_COLS, is_non_latin
)

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Leads Agent",
    page_icon="⚙️",
    layout="centered",
)

st.title("⚙️ Leads Agent")
st.caption("Upload a raw lead file → select project → download clean CSV")

# ── Sidebar: project selector ─────────────────────────────────────────────────
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
        help="Required only for files with non-Latin text (Korean, Arabic, Chinese, etc.). Get a free key at console.groq.com"
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
    help=f"Expected format: {fmt}"
)

if uploaded:
    file_bytes = uploaded.read()
    st.success(f"✅ File loaded: **{uploaded.name}** ({len(file_bytes):,} bytes)")

    # ── Detect non-Latin ─────────────────────────────────────────────────────
    with st.spinner("Scanning for non-Latin text..."):
        to_translate = detect_non_latin_fields(file_bytes, config)

    translated = {}

    if to_translate:
        st.warning(f"⚠️ Non-Latin text detected in {len(to_translate)} field(s). Translation needed.")
        with st.expander("Fields to translate", expanded=True):
            for field, items in to_translate.items():
                for idx, val in items.items():
                    st.write(f"**{field}** (row {idx}): `{val}`")

        if not api_key:
            st.error("Please enter your Groq API key in the sidebar to translate these fields. Get a free key at console.groq.com")
            st.stop()

        with st.spinner("Translating non-Latin fields via Claude..."):
            all_values = []
            index_map = []
            for field, items in to_translate.items():
                for idx, val in items.items():
                    all_values.append(val)
                    index_map.append((field, idx))

            try:
                resp = requests.post(
                    "https://api.groq.com/openai/v1/chat/completions",
                    headers={
                        "Authorization": f"Bearer {api_key}",
                        "Content-Type": "application/json",
                    },
                    json={
                        "model": "llama-3.3-70b-versatile",
                        "max_tokens": 1000,
                        "messages": [
                            {
                                "role": "system",
                                "content": (
                                    "You are a translation assistant. "
                                    "Translate any non-English or non-Latin script text to English. "
                                    "If already in English/Latin, return as-is. "
                                    "Return ONLY a JSON array of translated strings in the same order, "
                                    "no explanation, no markdown backticks."
                                )
                            },
                            {
                                "role": "user",
                                "content": f"Input: {json.dumps(all_values)}"
                            }
                        ]
                    },
                    timeout=30,
                )
                result = resp.json()
                text = result["choices"][0]["message"]["content"]
                clean = text.replace("```json", "").replace("```", "").strip()
                trans_list = json.loads(clean)

                for (field, idx), trans_val in zip(index_map, trans_list):
                    translated.setdefault(field, {})[idx] = trans_val

                st.success("✅ Translation complete")
                with st.expander("Translation results"):
                    for (field, idx), trans_val in zip(index_map, trans_list):
                        orig = to_translate[field][idx]
                        st.write(f"**{field}**: `{orig}` → `{trans_val}`")

            except Exception as e:
                st.error(f"Translation failed: {e}. Proceeding with original text.")

    # ── Process ───────────────────────────────────────────────────────────────
    with st.spinner("Processing leads..."):
        try:
            result_df = process(file_bytes, config, translated if translated else None)
        except Exception as e:
            st.error(f"Processing error: {e}")
            st.stop()

    # ── Stats ─────────────────────────────────────────────────────────────────
    col1, col2, col3 = st.columns(3)
    col1.metric("Unique contacts", len(result_df))
    col2.metric("LeadSource1", config.get("lead_source_1", "—"))
    col3.metric("LeadSource2", config.get("lead_source_2", "—"))

    # ── Preview ───────────────────────────────────────────────────────────────
    with st.expander("Preview output (first 5 rows)"):
        preview_cols = ["Email", "FirstName", "LastName", "Company", "City", "Country", "LeadSource1", "LeadSource2"]
        st.dataframe(result_df[[c for c in preview_cols if c in result_df.columns]].head())

    # ── Download ──────────────────────────────────────────────────────────────
    fname_base = uploaded.name.rsplit(".", 1)[0]
    out_name = f"{project_name.replace(' ', '_')}_{fname_base}_leads.csv"
    csv_bytes = to_csv_bytes(result_df)

    st.download_button(
        label="⬇️ Download CSV",
        data=csv_bytes,
        file_name=out_name,
        mime="text/csv",
        use_container_width=True,
    )

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
st.caption("To add a new project, add a config block to `configs.py` — no other code changes needed.")
