import streamlit as st
from openpyxl import load_workbook
from io import BytesIO
import re
import openpyxl

st.set_page_config(page_title="Named Range Cell Coordinates", layout="wide")
st.title("\U0001F4C2 Named Range Coordinate Extractor")
st.write("Upload one or more Excel files. For each named range, the app will display all cell coordinates in the form of [WorkbookName][SheetName]Cell[row][col], the associated formula or value, and a mapped reference using named ranges if applicable.")

uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.header(f"\U0001F4C4 File: {uploaded_file.name}")
        workbook_bytes = BytesIO(uploaded_file.read())
        wb = load_workbook(workbook_bytes, data_only=False)

        named_ranges_map = {}  # Structure: {(sheet_name, row, col): (name, r_offset, c_offset)}
        named_range_definitions = {}  # Structure: {name: (workbook, sheet, ref)}

        # Build named ranges map
        for name in wb.defined_names:
            dn_obj = wb.defined_names[name]
            if dn_obj.is_external or not dn_obj.attr_text:
                continue
            for sheet_name, ref in dn_obj.destinations:
                try:
                    ws = wb[sheet_name]
                    coord = ref.replace("$", "").split("!")[-1]
                    cell_range = ws[coord] if ":" in coord else [[ws[coord]]]
                    min_row = min(cell.row for row in cell_range for cell in row)
                    min_col = min(cell.column for row in cell_range for cell in row)
                    for row in cell_range:
                        for cell in row:
                            key = (sheet_name, cell.row, cell.column)
                            named_ranges_map[key] = (name, cell.row - min_row + 1, cell.column - min_col + 1)
                    named_range_definitions[name] = (uploaded_file.name, sheet_name, coord)
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

                            if isinstance(cell.value, str) and cell.value.startswith("="):
                                cell_content = cell.value.strip()
                            elif hasattr(cell.value, "text"):
                                cell_content = str(cell.value.text).strip()
                            elif cell.value is not None:
                                cell_content = f"[value] {cell.value}"
                            else:
                                cell_content = "(empty)"

                            # Reference formula mapping
                            def map_reference(m):
                                full_ref = m.group(0)
                                parts = full_ref.split("!")

                                if len(parts) == 2:
                                    sheet_ref, cell_ref = parts[0], parts[1]
                                else:
                                    sheet_ref, cell_ref = sheet_name, parts[0]  # default to current sheet

                                cell_ref = cell_ref.replace("$", "").upper()
                                match = re.match(r"([A-Z]+)([0-9]+)", cell_ref)
                                if not match:
                                    return full_ref

                                col_letter, row_number = match.groups()
                                row_num = int(row_number)
                                col_num = openpyxl.utils.column_index_from_string(col_letter)

                                key = (sheet_ref, row_num, col_num)
                                if key in named_ranges_map:
                                    nr_name, r_offset, c_offset = named_ranges_map[key]
                                    return f"{nr_name}[{r_offset}][{c_offset}]"

                                return f"[{uploaded_file.name}][{sheet_ref}]Cell[{row_num}][{col_num}]"

                            reference_formula = re.sub(r"(?:[A-Za-z0-9_]+!)?[A-Z]{1,3}[0-9]{1,7}", map_reference, cell_content) if isinstance(cell_content, str) else cell_content

                            entries.append(f"{cell_label} = {cell_content}\n → {reference_formula}")
                except Exception as e:
                    entries.append(f"Error accessing {ref}: {e}")

            workbook_name, sheet_name_for_range, ref_string = named_range_definitions.get(name, (uploaded_file.name, sheet_name, ref))
            with st.expander(f"\U0001F4CC Named Range: {name} [{workbook_name}][{sheet_name_for_range}][{ref_string}]"):
                st.write("**Cell Coordinates, Raw Formula/Value, and Reference Formula:**")
                st.code("\n".join(entries), language="text")
else:
    st.info("⬆️ Upload one or more `.xlsx` files to get started.")
