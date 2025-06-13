import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from io import BytesIO
import re
import os
from collections import defaultdict
import graphviz

st.set_page_config(page_title="Named Range Formula Remapper", layout="wide")
st.title("\U0001F4D8 Named Range Coordinates + Formula Remapping")

# --- Force expanders open and print all content nicely ---
st.markdown("""
    <style>
        @media print {
            html, body, [data-testid="stAppViewContainer"], [data-testid="stVerticalBlock"] {
                overflow: visible !important;
                height: auto !important;
            }
            details { display: block !important; }
        }
    </style>
    <script>
    window.addEventListener('load', function() {
        document.querySelectorAll('details').forEach(el => el.open = true);
    });
    </script>
""", unsafe_allow_html=True)

# --- Print Button ---
st.markdown("""
    <div style='text-align: right; margin-bottom: 1em;'>
        <button onclick="window.print()">🖨️ Print This Page</button>
    </div>
""", unsafe_allow_html=True)

# Allow manual mapping of external references like [1], [2], etc.
st.subheader("Manual Mapping for External References")
external_refs = {}
for i in range(1, 10):
    ref_key = f"[{i}]"
    workbook_name = st.text_input(f"Map external reference {ref_key} to workbook name (e.g., Mortality_Model_Inputs.xlsx)", key=ref_key)
    if workbook_name:
        external_refs[ref_key] = workbook_name

uploaded_files = st.file_uploader("\U0001F4C2 Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_named_cell_map = {}
    all_named_ref_info = {}
    file_display_names = {}
    named_ref_formulas = {}

    for uploaded_file in uploaded_files:
        display_name = uploaded_file.name
        file_display_names[display_name] = uploaded_file
        st.header(f"\U0001F4C4 File: `{display_name}`")
        wb = load_workbook(BytesIO(uploaded_file.read()), data_only=False)

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
                            all_named_cell_map[(display_name, sheet_name, r, c)] = (name, row_offset, col_offset)
                            coord_set.add((r, c))
                    all_named_ref_info[name] = (display_name, sheet_name, coord_set, min_row, min_col)
                except:
                    continue

    def remap_formula(formula, current_file, current_sheet):
        if not formula:
            return ""

        def cell_address(row, col):
            return f"{get_column_letter(col)}{row}"

        def remap_single_cell(ref, default_file, default_sheet):
            if "!" in ref:
                sheet_part, addr = ref.split("!")
                match = re.match(r"\\[(\\d+)\\]", sheet_part)
                if match:
                    external_ref = match.group(0)
                    external_file = external_refs.get(external_ref, external_ref)
                    return f"[{external_file}]{ref}"
                sheet_name = sheet_part
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

            key = (default_file, sheet_name, row, col)
            if key in all_named_cell_map:
                name, r_off, c_off = all_named_cell_map[key]
                return f"[{default_file}]{name}[{r_off}][{c_off}]"
            else:
                return f"[{default_file}]{sheet_name}!{addr}"

        def remap_range(ref, default_file, default_sheet):
            if ref.startswith("["):
                match = re.match(r"\\[(\\d+)\\]", ref)
                if match:
                    external_ref = match.group(0)
                    external_file = external_refs.get(external_ref, external_ref)
                    return f"[{external_file}]{ref}"

            if "!" in ref:
                sheet_name, addr = ref.split("!")
            else:
                sheet_name = default_sheet
                addr = ref

            addr = addr.replace("$", "").upper()
            if ":" not in addr:
                return remap_single_cell(ref, default_file, default_sheet)

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
                    key = (default_file, sheet_name, row, col)
                    if key in all_named_cell_map:
                        name, r_off, c_off = all_named_cell_map[key]
                        label_set.add(f"[{default_file}]{name}[{r_off}][{c_off}]")
                    else:
                        label_set.add(f"[{default_file}]{sheet_name}!{cell_address(row, col)}")
            return ", ".join(sorted(label_set))

        pattern = r"(?<![A-Za-z0-9_])(?:\\[[^\\]]+\\])?[A-Za-z0-9_]+!\\$?[A-Z]{1,3}\\$?[0-9]{1,7}(?::\\$?[A-Z]{1,3}\\$?[0-9]{1,7})?|(?<![A-Za-z0-9_])\\$?[A-Z]{1,3}\\$?[0-9]{1,7}(?::\\$?[A-Z]{1,3}\\$?[0-9]{1,7})?"
        matches = list(re.finditer(pattern, formula))
        replaced_formula = formula
        offset = 0
        for match in matches:
            raw = match.group(0)
            remapped = remap_range(raw, current_file, current_sheet)
            start, end = match.start() + offset, match.end() + offset
            replaced_formula = replaced_formula[:start] + remapped + replaced_formula[end:]
            offset += len(remapped) - len(raw)
        return replaced_formula

    # Rest of the app logic (remapping and rendering) stays unchanged...
