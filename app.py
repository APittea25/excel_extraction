import streamlit as st
from openpyxl import load_workbook
import io
import pandas as pd

st.title("📘 Excel Named References Viewer")

uploaded_files = st.file_uploader(
    "Upload one or more Excel files (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

def extract_named_references(wb):
    named_refs = []

    for name in wb.defined_names:
        dn = wb.defined_names[name]
        if dn.is_external or not dn.attr_text:
            continue

        for sheet_title, ref in dn.destinations:
            named_refs.append({
                "Name": name,
                "Sheet": sheet_title,
                "Reference": ref
            })

    return pd.DataFrame(named_refs)

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.subheader(f"📄 {uploaded_file.name}")
        try:
            wb = load_workbook(io.BytesIO(uploaded_file.read()), data_only=False)
            df = extract_named_references(wb)
            if not df.empty:
                st.dataframe(df)
            else:
                st.info("No named references found.")
        except Exception as e:
            st.error(f"❌ Error reading {uploaded_file.name}: {e}")
