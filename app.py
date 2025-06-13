import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from io import BytesIO
import re

st.set_page_config(page_title="Named Range Formula Remapper", layout="wide")
st.title("📘 Named Range Coordinates + Formula Remapping")

uploaded_files = st.file_uploader("📂 Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.header(f"📄 File: `{uploaded_file.name}`")
        wb = load_workbook(BytesIO(uploaded_file.read()), data_only=False)

        # Map of: (sheet, row, col) -> (named_ref, row_offset, col_offset)
        named_cell_map = {}

        # Named reference metadata: name -> (sheet, set of (row, col), top_left_row, top_left_col)
        named_ref_info = {}

        # Build named cell lookup
        for name in wb.defined_names:
            dn = wb.defined_names[name]
            if dn.is_external or not dn.attr_text:
                continue
            for sheet_name, ref in dn.destinations:
                try:
                    ws = wb[sheet_name]
                    ref_clean = ref.replace("$", "").split("!")[-1]
                    cells = ws[ref_clean] if ":" in ref_clean else [[ws[ref_clean]]]

                    min_row = min(cell.row for row in cells for cell in row)
                    min_col = min(cell.column for row in cells for cell in row)

                    coord_set = set()
                    for row in cells:
                        for cell in row:
                            r, c = cell.row, cell.column
                            row_offset = r - min_row + 1
                            col_offset = c - min_col + 1
                            named_cell_map[(sheet_name, r, c)] = (name, row_offset, col_offset)
                            coord_set.add((r, c))
                    named_ref_info[name] = (sheet_name, coord_set, min_row, min_col)
                except:
                    continue

        # Remapping function
        def remap_formula(formula, current_sheet):
            def replace_match(m):
                raw = m.group(0)
                if "!" in raw:
                    sheet_part, addr = raw.split("!")
                    target_sheet = sheet_part
                else:
                    target_sheet = current_sheet
                    addr = raw

                addr = addr.replace("$", "").upper()
                match = re.match(r"([A-Z]+)([0-9]+)", addr)
                if not match:
                    return raw

                col_str, row_str = match.groups()
                row = int(row_str)
                col = column_index_from_string(col_str)

                key = (target_sheet, row, col)
                if key in named_cell_map:
                    name, r_off, c_off = named_cell_map[key]
                    return f"{name}[{r_off}][{c_off}]"
                else:
                    return raw  # not inside any named reference

            # Only change plain cell refs or Sheet!Cell
            return re.sub(r"(?:[A-Za-z0-9_]+!)?[A-Z]{1,3}[0-9]{1,7}", replace_match, formula)

        # Display remapped formulas
        for name in wb.defined_names:
            dn = wb.defined_names[name]
            if dn.is_external or not dn.attr_text:
                continue

            entries = []

            for sheet_name, ref in dn.destinations:
                try:
                    ws = wb[sheet_name]
                    ref_clean = ref.replace("$", "").split("!")[-1]
                    cell_range = ws[ref_clean] if ":" in ref_clean else [[ws[ref_clean]]]

                    min_row = min(cell.row for row in cell_range for cell in row)
                    min_col = min(cell.column for row in cell_range for cell in row)

                    for row in cell_range:
                        for cell in row:
                            row_offset = cell.row - min_row + 1
                            col_offset = cell.column - min_col + 1
                            label = f"{name}[{row_offset}][{col_offset}]"

                            # Determine raw content
                            if isinstance(cell.value, str) and cell.value.startswith("="):
                                formula = cell.value.strip()
                                remapped = remap_formula(formula, sheet_name)
                            elif cell.value is not None:
                                formula = f"[value] {cell.value}"
                                remapped = formula
                            else:
                                formula = "(empty)"
                                remapped = formula

                            entries.append(f"{label} = {formula}\n → {remapped}")
                except Exception as e:
                    entries.append(f"❌ Error accessing `{ref}`: {e}")

            with st.expander(f"📌 Named Range: `{name}` → {ref}"):
                st.code("\n".join(entries), language="text")
else:
    st.info("⬆️ Upload `.xlsx` files to begin.")
