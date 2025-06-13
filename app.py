import streamlit as st
from openpyxl import load_workbook
from io import BytesIO
import re

st.set_page_config(page_title="Named Range Inspector", layout="wide")
st.title("📊 Excel Named Range Inspector")
st.write("Upload one or more Excel files to inspect named ranges, their location, and formulas across all referenced cells.")

uploaded_files = st.file_uploader("Upload Excel files", type=['xlsx'], accept_multiple_files=True)

# Session state to toggle expand/collapse
if "expand_all" not in st.session_state:
    st.session_state.expand_all = False

def toggle_expand():
    st.session_state.expand_all = not st.session_state.expand_all
    st.experimental_rerun()

if uploaded_files:
    st.button(
        "Expand All" if not st.session_state.expand_all else "Collapse All",
        on_click=toggle_expand,
        help="Toggle between expanding or collapsing all named ranges"
    )

    def extract_named_ranges(file, filename):
        wb = load_workbook(filename=BytesIO(file.read()), data_only=False)
        result = []

        for name in wb.defined_names:
            dn = wb.defined_names[name]

            if dn.is_external or not dn.attr_text:
                continue

            destinations = list(dn.destinations)

            for sheet_name, ref in destinations:
                try:
                    ws = wb[sheet_name]
                    coord = ref.replace("$", "").split("!")[-1]
                    formulas = []

                    # Handle single cell or range
                    try:
                        cell_range = ws[coord]
                        if not isinstance(cell_range, (tuple, list)):
                            cell_range = [[cell_range]]
                    except Exception as e:
                        result.append({
                            "Named Range": name,
                            "File": filename,
                            "Sheet": sheet_name,
                            "Range": coord,
                            "Formulas": [f"Error reading cells: {str(e)}"]
                        })
                        continue

                    for row in cell_range:
                        for cell in row:
                            raw_formula = None

                            if isinstance(cell.value, str) and cell.value.startswith("="):
                                raw_formula = cell.value.strip()
                            elif hasattr(cell.value, "text"):
                                raw_formula = str(cell.value.text).strip()

                            if raw_formula:
                                formulas.append(raw_formula)
                            elif cell.value is not None:
                                formulas.append(f"[value] {cell.value}")

                    result.append({
                        "Named Range": name,
                        "File": filename,
                        "Sheet": sheet_name,
                        "Range": coord,
                        "Formulas": formulas if formulas else ["(No formulas or values found)"]
                    })

                except Exception as e:
                    result.append({
                        "Named Range": name,
                        "File": filename,
                        "Sheet": sheet_name,
                        "Range": ref,
                        "Formulas": [f"Error accessing range: {str(e)}"]
                    })

        return result

    for uploaded_file in uploaded_files:
        st.header(f"🔍 File: {uploaded_file.name}")
        results = extract_named_ranges(uploaded_file, uploaded_file.name)

        for item in results:
            with st.expander(f"📌 Named Range: {item['Named Range']}", expanded=st.session_state.expand_all):
                st.write(f"**Sheet:** {item['Sheet']}")
                st.write(f"**Range:** {item['Range']}")
                st.write("**Formulas / Values:**")
                st.code("\n".join(item["Formulas"]), language="excel")
else:
    st.info("Upload one or more .xlsx files to begin analysis.")
