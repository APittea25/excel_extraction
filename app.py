import streamlit as st
from openpyxl import load_workbook
import pandas as pd
import io

st.title("📘 Excel Named References + Formulas Viewer")

uploaded_files = st.file_uploader(
    "Upload one or more Excel files (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True
)

def extract_formulas_in_range(wb, sheet_name, ref):
    formulas = []
    try:
        sheet = wb[sheet_name]
        cells = sheet[ref]  # could be a single cell or range

        if isinstance(cells, tuple):
            for row in cells:
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        formulas.append(cell.value.strip())
        else:
            # Single cell
            cell = cells
            if isinstance(cell.value, str) and cell.value.startswith("="):
                formulas.append(cell.value.strip())
    except Exception as e:
        formulas.append(f"(Error reading range: {e})")
    return formulas

def extract_named_references(wb, workbook_name):
    named_refs = []

    for name in wb.defined_names:
        dn = wb.defined_names[name]
        if dn.is_external or not dn.attr_text:
            continue

        for sheet_title, ref in dn.destinations:
            formulas = extract_formulas_in_range(wb, sheet_title, ref)
            named_refs.append({
                "Workbook": workbook_name,
                "Name": name,
                "Sheet": sheet_title,
                "Reference": ref,
                "Formulas": "\n".join(formulas) if formulas else "(No formulas)"
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
        st.subheader("📋 Named References with Formulas")
        st.dataframe(df)
    else:
        st.info("No named references found.")
else:
    st.info("⬆️ Upload `.xlsx` files to begin.")
