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

        # Map named ranges to coordinate sets
        named_ranges_map = {}
        for defined_name in wb.defined_names:
            dn_obj = wb.defined_names[defined_name]
            if dn_obj.is_external or not dn_obj.attr_text:
                continue
            for sheet_name, ref in dn_obj.destinations:
                try:
                    ws = wb[sheet_name]
                    coord = ref.replace("$", "").split("!")[-1]
                    cell_range = ws[coord] if ":" in coord else [[ws[coord]]]
                    coords = {(cell.coordinate): (ri+1, ci+1) for ri, row in enumerate(cell_range) for ci, cell in enumerate(row)}
                    named_ranges_map[defined_name] = {"sheet": sheet_name, "cells": coords}
                except:
                    continue

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
                        cell_range = ws[coord] if ":" in coord else [[ws[coord]]]
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

                                # Attempt replacement of any direct cell reference
                                def replace_match(m):
                                    cell_ref = m.group(0)
                                    for nr_name, nr_data in named_ranges_map.items():
                                        if nr_data["sheet"] == sheet_name and cell_ref in nr_data["cells"]:
                                            r, c = nr_data["cells"][cell_ref]
                                            return f"{nr_name}[{r}][{c}]"
                                    return cell_ref

                                raw_formula = re.sub(r"[A-Z]+[0-9]+", replace_match, raw_formula)
                                formulas.append(raw_formula)

                            elif hasattr(cell.value, "text"):
                                raw_formula = str(cell.value.text).strip()
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
