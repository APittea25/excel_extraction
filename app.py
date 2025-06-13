import streamlit as st
from openpyxl import load_workbook
from io import BytesIO
import re
import openpyxl

st.set_page_config(page_title="Named Range Inspector", layout="wide")
st.title("\U0001F4CA Excel Named Range Inspector")
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

        # Build a mapping of named ranges to their cells and relative positions
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
                    min_row = min(cell.row for row in cell_range for cell in row)
                    min_col = min(cell.column for row in cell_range for cell in row)
                    # Store coordinate to relative row/col in the named range
                    coords = {
                        (cell.coordinate): (cell.row - min_row + 1, cell.column - min_col + 1)
                        for row in cell_range for cell in row
                    }
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

                    # Safely get cell range
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

                            # Handle string-based formulas
                            if isinstance(cell.value, str) and cell.value.startswith("="):
                                raw_formula = cell.value.strip()

                                # This function replaces literal cell refs with named range indexes
                                def replace_match(m):
                                    cell_ref = m.group(0).replace("$", "").upper()
                                    try:
                                        col_letters = re.match(r"[A-Z]+", cell_ref).group()
                                        row_numbers = re.match(r"[A-Z]+([0-9]+)", cell_ref).group(1)
                                        ref_col = openpyxl.utils.column_index_from_string(col_letters)
                                        ref_row = int(row_numbers)
                                    except Exception:
                                        return m.group(0)

                                    for nr_name, nr_data in named_ranges_map.items():
                                        # First try exact match
                                        for coord, (r_offset, c_offset) in nr_data["cells"].items():
                                            col_str, row = openpyxl.utils.cell.coordinate_from_string(coord)
                                            col = openpyxl.utils.column_index_from_string(col_str)
                                            if row == ref_row and col == ref_col:
                                                return f"{nr_name}[{r_offset}][{c_offset}]"
                                        # Fallback bounding box check
                                        try:
                                            coords = list(nr_data["cells"].keys())
                                            min_r = min(openpyxl.utils.cell.coordinate_to_tuple(c)[0] for c in coords)
                                            max_r = max(openpyxl.utils.cell.coordinate_to_tuple(c)[0] for c in coords)
                                            min_c = min(openpyxl.utils.cell.coordinate_to_tuple(c)[1] for c in coords)
                                            max_c = max(openpyxl.utils.cell.coordinate_to_tuple(c)[1] for c in coords)
                                            if min_r <= ref_row <= max_r and min_c <= ref_col <= max_c:
                                                rel_r = ref_row - min_r + 1
                                                rel_c = ref_col - min_c + 1
                                                return f"{nr_name}[{rel_r}][{rel_c}]"
                                        except:
                                            continue

                                    return m.group(0)

                                # Substitute all simple cell references with mapping logic
                                raw_formula = re.sub(r"\b[A-Z]{1,3}[0-9]{1,7}\b", replace_match, raw_formula)
                                formulas.append(raw_formula)

                            # Handle other formula types or static values
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
        st.header(f"\U0001F50D File: {uploaded_file.name}")
        results = extract_named_ranges(uploaded_file, uploaded_file.name)

        for item in results:
            with st.expander(f"\U0001F4CC Named Range: {item['Named Range']}", expanded=st.session_state.expand_all):
                st.write(f"**Sheet:** {item['Sheet']}")
                st.write(f"**Range:** {item['Range']}")
                st.write("**Formulas / Values:**")
                st.code("\n".join(item["Formulas"]), language="excel")
else:
    st.info("Upload one or more .xlsx files to begin analysis.")
