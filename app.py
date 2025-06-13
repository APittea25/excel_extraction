import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.title("Named Ranges Extractor from Excel Files")

uploaded_files = st.file_uploader(
    "Upload one or more Excel files", 
    type=["xlsx"], 
    accept_multiple_files=True
)

def extract_named_ranges(file):
    wb = load_workbook(filename=BytesIO(file.read()), data_only=True)
    named_ranges = []
    for name in wb.defined_names.definedName:
        name_info = {
            "Name": name.name,
            "Scope": name.localSheetId if name.localSheetId is not None else "Workbook",
            "Refers To": name.attr_text
        }
        named_ranges.append(name_info)
    return pd.DataFrame(named_ranges)

if uploaded_files:
    for file in uploaded_files:
        st.subheader(f"Named Ranges in: {file.name}")
        try:
            df = extract_named_ranges(file)
            if not df.empty:
                st.dataframe(df)
            else:
                st.info("No named ranges found.")
        except Exception as e:
            st.error(f"Failed to process {file.name}: {e}")
