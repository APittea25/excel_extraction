import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from io import BytesIO
import re
from collections import defaultdict
import graphviz

st.set_page_config(page_title="Named Range Formula Remapper", layout="wide")
st.title("📘 Named Range Coordinates + Formula Remapping")

# Expand all expanders on load
st.markdown("""
    <script>
    window.addEventListener('load', function() {
        document.querySelectorAll('details').forEach(el => el.open = true);
    });
    </script>
""", unsafe_allow_html=True)

# Manual reference mapping
st.subheader("Manual Mapping for External References")
external_refs = {}
for i in range(1, 10):
    key = f"[{i}]"
    val = st.text_input(f"Map {key} to workbook name", key=key)
    if val:
        external_refs[key] = val

uploaded_files = st.file_uploader("📁 Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_named_cell_map = {}
    all_named_ref_info = {}
    file_display_names = {}
    named_ref_formulas = {}

    for uploaded_file in uploaded_files:
        display_name = uploaded_file.name
        file_display_names[display_name] = uploaded_file
        st.header(f"📄 File: `{display_name}`")
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
                            all_named_cell_map[(display_name, sheet_name, r, c)] = (name, r - min_row + 1, c - min_col + 1)
                            coord_set.add((r, c))
                    all_named_ref_info[name] = (display_name, sheet_name, coord_set, min_row, min_col)
                except:
                    continue

    def remap_formula(formula, current_file, current_sheet):
        if not formula:
            return ""
        def cell_address(row, col):
            return f"{get_column_letter(col)}{row}"

        pattern = r"(?:\[[^\]]+\])?[A-Za-z0-9_]+!\$?[A-Z]{1,3}\$?[0-9]+(?::\$?[A-Z]{1,3}\$?[0-9]+)?"
        parts = re.finditer(pattern, formula)
        result = formula
        offset = 0

        for part in parts:
            raw = part.group(0)
            ref = raw.split("!")[-1].replace("$", "")
            if ":" in ref:
                result = result[:part.start() + offset] + f"[{current_file}]{current_sheet}!{ref}" + result[part.end() + offset:]
                offset += len(f"[{current_file}]{current_sheet}!{ref}") - len(raw)
            else:
                col = column_index_from_string(re.findall(r"[A-Z]+", ref)[0])
                row = int(re.findall(r"[0-9]+", ref)[0])
                key = (current_file, current_sheet, row, col)
                if key in all_named_cell_map:
                    name, ro, co = all_named_cell_map[key]
                    new_ref = f"[{current_file}]{name}[{ro}][{co}]"
                    result = result[:part.start() + offset] + new_ref + result[part.end() + offset:]
                    offset += len(new_ref) - len(raw)
        return result

    reverse_refs = defaultdict(list)

    for name, (file_name, sheet_name, coords, min_row, min_col) in all_named_ref_info.items():
        try:
            wb = load_workbook(BytesIO(file_display_names[file_name].getvalue()), data_only=False)
            ws = wb[sheet_name]
            ref_cells = [[ws.cell(r, c) for c in range(min(c for (_, c) in coords), max(c for (_, c) in coords)+1)]
                         for r in range(min(r for (r, _) in coords), max(r for (r, _) in coords)+1)]
            entries = []
            formulas = []

            for row in ref_cells:
                for cell in row:
                    r_offset = cell.row - min_row + 1
                    c_offset = cell.column - min_col + 1
                    label = f"{name}[{r_offset}][{c_offset}]"
                    formula = str(cell.value).strip() if cell.value and str(cell.value).startswith("=") else ""
                    remapped = remap_formula(formula, file_name, sheet_name) if formula else str(cell.value)
                    if remapped:
                        formulas.append(remapped)
                    entries.append(f"{label} = {formula}\n → {remapped}")
            named_ref_formulas[name] = formulas
            with st.expander(f"📌 `{name}` → `{sheet_name}` in `{file_name}`"):
                st.code("\n".join(entries), language="text")
        except Exception as e:
            st.error(f"❌ Error processing {name}: {e}")

    # ✅ Efficient dependency analysis
    st.subheader("🔗 Dependency Graph")
    dot = graphviz.Digraph()
    dot.attr(compound='true', rankdir='LR')

    group_by_file = defaultdict(list)
    for name, (fname, sheet, *_rest) in all_named_ref_info.items():
        group_by_file[fname].append((name, sheet))

    all_names = set(named_ref_formulas.keys())

    for idx, (file, nodes) in enumerate(group_by_file.items()):
        with dot.subgraph(name=f"cluster_{idx}") as c:
            c.attr(label=file, style='filled', color='lightgrey')
            for name, sheet in nodes:
                c.node(name, color="blue" if "Sheet" in sheet else "black")

    for tgt, formulas in named_ref_formulas.items():
        formula_text = " ".join(formulas)
        for src in all_names:
            if src != tgt and src in formula_text:
                dot.edge(src, tgt)

    st.graphviz_chart(dot)
else:
    st.info("⬆️ Upload one or more `.xlsx` files to begin.")
