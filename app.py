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

def extract_named_references(wb, workbook_name):
    named_refs = []

    for name in wb.defined_names:
        dn = wb.defined_names[name]
        if dn.is_external or not dn.attr_text:
            continue

        for sheet_title, ref in dn.destinations:
            named_refs.append({
                "Workbook": workbook_name,
                "Sheet": sheet_title,
                "Reference": ref,
                "Name": name
            })

    return named_refs

all_refs = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        try:
            wb = load_workbook(io.BytesIO(uploaded_file.read()), data_only=False)
            refs = extract_named_references(wb, uploaded_file.name)
            all_refs.extend(refs)
        except Exception as e:
            st.error(f"❌ Error reading {uploaded_file.name}: {e}")

    if all_refs:
        df = pd.DataFrame(all_refs)
        st.subheader("📋 Named References Summary")
        st.dataframe(df)
    else:
        st.info("No named references found in uploaded files.")
else:
    st.info("⬆️ Upload one or more `.xlsx` files to begin.")
