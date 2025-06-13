import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from io import BytesIO
import re

st.set_page_config(page_title="Named Range Formula Remapper", layout="wide")
st.title("📘 Named Range Coordinates + Formula Remapping")

uploaded_files = st.file_uploader("📂 Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.header(f"📄 File: `{uploaded_file.name}`")
        wb = load_workbook(BytesIO(uploaded_file.read()), data_only=False)

        named_cell_map = {}
        named_ref_info = {}

        # Step 1: Map all named references to cell positions
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

        # Step 2: Remapping logic
        def remap_formula(formula, current_sheet):
            if not formula:
                return ""
            formula = re.sub(r"'[^']+\\.xlsx'!", "", formula)
            formula = re.sub(r"\\[[^\\]]+\\][^!]*!", "", formula)

            def cell_address(row, col):
                return f"{get_column_letter(col)}{row}"

            def remap_single_cell(ref, default_sheet):
                if "!" in ref:
                    sheet_name, addr = ref.split("!")
                else:
                    sheet_name = default_sheet
                    addr = ref
                addr = addr.replace("$", "").upper()
                match = re.match(r"([A-Z]+)([0-9]+)", addr)
                if not match:
                    return ref
                col_str, row_str = match.groups()
                row = int(row_str)
                col = column_index_from_string(col_str)
                key = (sheet_name, row, col)
                if key in named_cell_map:
                    name, r_off, c_off = named_cell_map[key]
                    return f"{name}[{r_off}][{c_off}]"
                else:
                    return ref

            def remap_range(ref, default_sheet):
                if "!" in ref:
                    sheet_name, addr = ref.split("!")
                else:
                    sheet_name = default_sheet
                    addr = ref
                addr = addr.replace("$", "").upper()
                if ":" not in addr:
                    return remap_single_cell(ref, default_sheet)

                start, end = addr.split(":")
                m1 = re.match(r"([A-Z]+)([0-9]+)", start)
                m2 = re.match(r"([A-Z]+)([0-9]+)", end)
                if not m1 or not m2:
                    return ref
                start_col = column_index_from_string(m1.group(1))
                start_row = int(m1.group(2))
                end_col = column_index_from_string(m2.group(1))
                end_row = int(m2.group(2))

                label_set = set()
                for row in range(start_row, end_row + 1):
                    for col in range(start_col, end_col + 1):
                        key = (sheet_name, row, col)
                        if key in named_cell_map:
                            name, r_off, c_off = named_cell_map[key]
                            label_set.add(f"{name}[{r_off}][{c_off}]")
                        else:
                            label_set.add(f"{sheet_name}!{cell_address(row, col)}")
                return ", ".join(sorted(label_set)) if label_set else ref

            pattern = r"(?<![A-Za-z0-9_])(?:[A-Za-z0-9_]+!)?\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?"
            matches = list(re.finditer(pattern, formula))
            replaced_formula = formula
            offset = 0
            for match in matches:
                raw = match.group(0)
                remapped = remap_range(raw, current_sheet)
                start, end = match.start() + offset, match.end() + offset
                replaced_formula = replaced_formula[:start] + remapped + replaced_formula[end:]
                offset += len(remapped) - len(raw)
            return replaced_formula

        # Step 3: Display remapped formulas per named reference
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

                            try:
                                raw_formula = None
                                if cell.data_type == 'f':
                                    raw_formula = str(cell.value)
                                elif hasattr(cell, 'formula') and hasattr(cell.formula, 'text'):
                                    raw_formula = str(cell.formula.text)

                                if raw_formula and raw_formula.startswith("="):
                                    formula = raw_formula.strip()
                                    remapped = remap_formula(formula, sheet_name)
                                elif cell.value is not None:
                                    formula = f"[value] {str(cell.value)}"
                                    remapped = formula
                                else:
                                    formula = "(empty)"
                                    remapped = formula
                            except Exception as e:
                                formula = f"[error reading cell: {e}]"
                                remapped = formula

                            entries.append(f"{label} = {formula}\n → {remapped}")
                except Exception as e:
                    entries.append(f"❌ Error accessing `{ref}`: {e}")

            with st.expander(f"📌 Named Range: `{name}` → {ref}"):
                st.code("\n".join(entries), language="text")
else:
    st.info("⬆️ Upload one or more `.xlsx` files to begin.")
