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

                            if isinstance(cell.value, str) and cell.value.startswith("="):
                                cell_content = cell.value.strip()
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
