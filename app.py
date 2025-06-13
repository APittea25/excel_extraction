import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from io import BytesIO
import re

st.set_page_config(page_title="Named Range Formula Explorer", layout="wide")
st.title("📘 Named Range Cell Coordinates & Formulas")
st.write("Upload one or more Excel files. For each named range, this app will display every cell within the range, its coordinates, raw formula/value, and a mapped version using other named ranges if relevant.")

uploaded_files = st.file_uploader("📂 Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.header(f"📄 File: `{uploaded_file.name}`")
        wb = load_workbook(BytesIO(uploaded_file.read()), data_only=False)

        named_range_cells = {}  # (sheet, row, col): (name, row_offset, col_offset)
        named_range_meta = {}   # name: (workbook, sheet, ref_string, set of (row, col))

        # First pass: Build named range coordinate maps
        for name in wb.defined_names:
            dn = wb.defined_names[name]
            if dn.is_external or not dn.attr_text:
                continue
            for sheet_name, ref in dn.destinations:
                try:
                    ws = wb[sheet_name]
                    ref_clean = ref.replace("$", "").split("!")[-1]
                    cell_range = ws[ref_clean] if ":" in ref_clean else [[ws[ref_clean]]]

                    min_row = min(cell.row for row in cell_range for cell in row)
                    min_col = min(cell.column for row in cell_range for cell in row)

                    coord_set = set()
                    for row in cell_range:
                        for cell in row:
                            key = (sheet_name, cell.row, cell.column)
                            named_range_cells[key] = (name, cell.row - min_row + 1, cell.column - min_col + 1)
                            coord_set.add((cell.row, cell.column))

                    named_range_meta[name] = (uploaded_file.name, sheet_name, ref_clean, coord_set)
                except Exception:
                    continue

        # Second pass: Display each named range and its cell contents
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
                            cell_label = f"[{uploaded_file.name}][{sheet_name}]Cell[{row_offset}][{col_offset}]"

                            # Raw content extraction
                            if isinstance(cell.value, str) and cell.value.startswith("="):
                                content = cell.value.strip()
                            elif hasattr(cell.value, "text"):
                                content = str(cell.value.text).strip()
                            elif cell.value is not None:
                                content = f"[value] {cell.value}"
                            else:
                                content = "(empty)"

                            # Reference remapping
                            def map_reference(match):
                                ref_text = match.group(0)
                                parts = ref_text.split("!")
                                if len(parts) == 2:
                                    ref_sheet, cell_addr = parts
                                else:
                                    ref_sheet, cell_addr = sheet_name, parts[0]

                                cell_addr = cell_addr.replace("$", "").upper()
                                match_coords = re.match(r"([A-Z]+)(\d+)", cell_addr)
                                if not match_coords:
                                    return ref_text

                                col_str, row_str = match_coords.groups()
                                row = int(row_str)
                                col = column_index_from_string(col_str)

                                key = (ref_sheet, row, col)
                                for nr, (_, nr_sheet, _, cell_set) in named_range_meta.items():
                                    if ref_sheet == nr_sheet and (row, col) in cell_set:
                                        if cell_set == {(row, col)}:
                                            return f"{nr}"
                                        return f"{nr}[{row}][{col}]"
                                return f"[{uploaded_file.name}][{ref_sheet}]Cell[{row}][{col}]"

                            mapped_content = re.sub(
                                r"(?:[A-Za-z0-9_]+!)?[A-Z]{1,3}[0-9]{1,7}",
                                map_reference,
                                content
                            ) if isinstance(content, str) else content

                            entries.append(f"{cell_label} = {content}\n → {mapped_content}")

                except Exception as e:
                    entries.append(f"❌ Error accessing `{ref}`: {e}")

            # Display result for this named range
            workbook_name, sheet_meta, ref_meta, _ = named_range_meta.get(name, (uploaded_file.name, sheet_name, ref, set()))
            with st.expander(f"📌 Named Range: `{name}` → [{workbook_name}][{sheet_meta}][{ref_meta}]"):
                st.write("**Coordinates, Formula/Value, and Mapped Reference**")
                st.code("\n".join(entries), language="text")
else:
    st.info("⬆️ Upload `.xlsx` files to begin.")
