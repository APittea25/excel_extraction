import streamlit as st
from openpyxl import load_workbook
from openpyxl.workbook.defined_name import DefinedName
from io import BytesIO

st.set_page_config(page_title="Named Range Inspector", layout="wide")
st.title("📊 Excel Named Range Inspector")
st.write("Upload one or more Excel files to inspect named ranges, their location, and formulas.")

uploaded_files = st.file_uploader("Upload Excel files", type=['xlsx'], accept_multiple_files=True)

def extract_named_ranges(file, filename):
    wb = load_workbook(filename=BytesIO(file.read()), data_only=False)
    result = []

    for name in wb.defined_names.defined_names:
        defined_name = wb.defined_names[name]

        # Skip if not referring to a range (e.g., constant)
        if not isinstance(defined_name, DefinedName):
            continue

        destinations = list(defined_name.destinations)

        for sheet_name, cell_range in destinations:
            try:
                ws = wb[sheet_name]
                formulas = []

                for row in ws[cell_range]:
                    for cell in row:
                        if cell.data_type == 'f':
                            formulas.append(cell.value)
                        elif cell.value is not None:
                            formulas.append(f"[value] {cell.value}")

                result.append({
                    "Named Range": name,
                    "File": filename,
                    "Sheet": sheet_name,
                    "Range": cell_range,
                    "Formulas": formulas if formulas else ["(No formulas or values found)"]
                })
            except Exception as e:
                result.append({
                    "Named Range": name,
                    "File": filename,
                    "Sheet": sheet_name,
                    "Range": cell_range,
                    "Formulas": [f"Error accessing range: {str(e)}"]
                })

    return result

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.header(f"🔍 File: {uploaded_file.name}")
        results = extract_named_ranges(uploaded_file, uploaded_file.name)

        for item in results:
            with st.expander(f"📌 Named Range: {item['Named Range']}"):
                st.write(f"**Sheet:** {item['Sheet']}")
                st.write(f"**Range:** {item['Range']}")
                st.write("**Formulas / Values:**")
                st.code("\n".join(item["Formulas"]), language="excel")
else:
    st.info("Upload one or more .xlsx files to begin analysis.")
