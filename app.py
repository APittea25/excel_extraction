import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Named Range Cell Viewer", layout="wide")
st.title("📘 Named Range Cell Coordinates & Formulas")
st.write("Upload one or more Excel files. This app will list all cells within each named range along with their coordinate and formula or value.")

uploaded_files = st.file_uploader("📂 Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.header(f"📄 File: `{uploaded_file.name}`")
        wb = load_workbook(BytesIO(uploaded_file.read()), data_only=False)

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

                            if isinstance(cell.value, str) and cell.value.startswith("="):
                                content = cell.value.strip()
                            elif hasattr(cell.value, "text"):
                                content = str(cell.value.text).strip()
                            elif cell.value is not None:
                                content = f"[value] {cell.value}"
                            else:
                                content = "(empty)"

                            entries.append(f"{cell_label} = {content}")

                except Exception as e:
                    entries.append(f"❌ Error accessing `{ref}`: {e}")

            with st.expander(f"📌 Named Range: `{name}` in `{sheet_name}` → {ref}"):
                st.write("**Coordinates and Formula/Value**")
                st.code("\n".join(entries), language="text")
else:
    st.info("⬆️ Upload `.xlsx` files to begin.")
