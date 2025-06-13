import streamlit as st
from openpyxl import load_workbook
from io import BytesIO
import re
import openpyxl

st.set_page_config(page_title="Named Range Cell Coordinates", layout="wide")
st.title("\U0001F4C2 Named Range Coordinate Extractor")
st.write("Upload one or more Excel files. For each named range, the app will display all cell coordinates in the form of [WorkbookName][SheetName]Cell[row][col] and the associated formula or value.")

uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.header(f"\U0001F4C4 File: {uploaded_file.name}")
        workbook_bytes = BytesIO(uploaded_file.read())
        wb = load_workbook(workbook_bytes, data_only=False)

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
            entries = []

            for sheet_name, ref in destinations:
                try:
                    ws = wb[sheet_name]
                    coord = ref.replace("$", "").split("!")[-1]
                    cell_range = ws[coord] if ":" in coord else [[ws[coord]]]
                    min_row = min(cell.row for row in cell_range for cell in row)
                    min_col = min(cell.column for row in cell_range for cell in row)

                    for row in cell_range:
                        for cell in row:
                            row_offset = cell.row - min_row + 1
                            col_offset = cell.column - min_col + 1
                            cell_label = f"[{uploaded_file.name}][{sheet_name}]Cell[{row_offset}][{col_offset}]"
                            raw_formula = None

                            # Handle string-based formulas
                            if isinstance(cell.value, str) and cell.value.startswith("="):
                                raw_formula = cell.value.strip()

                                # Replace cell references with named range indexes
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
                                        if nr_data["sheet"] != sheet_name:
                                            continue
                                        for coord, (r_offset, c_offset) in nr_data["cells"].items():
                                            col_str, row = openpyxl.utils.cell.coordinate_from_string(coord)
                                            col = openpyxl.utils.column_index_from_string(col_str)
                                            if row == ref_row and col == ref_col:
                                                return f"{nr_name}[{r_offset}][{c_offset}]"

                                    return m.group(0)

                                raw_formula = re.sub(r"\b[A-Z]{1,3}[0-9]{1,7}\b", replace_match, raw_formula)
                                cell_content = raw_formula

                            elif hasattr(cell.value, "text"):
                                cell_content = str(cell.value.text).strip()
                            elif cell.value is not None:
                                cell_content = f"[value] {cell.value}"
                            else:
                                cell_content = "(empty)"

                            entries.append(f"{cell_label} = {cell_content}")
                except Exception as e:
                    entries.append(f"Error accessing {ref}: {e}")

            with st.expander(f"\U0001F4CC Named Range: {name}"):
                st.write("**Cell Coordinates and Values/Formulas:**")
                st.code("\n".join(entries), language="text")
else:
    st.info("⬆️ Upload one or more `.xlsx` files to get started.")
