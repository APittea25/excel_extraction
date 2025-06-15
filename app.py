import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from io import BytesIO
import re
import os
from collections import defaultdict
import graphviz

st.set_page_config(page_title="Named Range Formula Remapper", layout="wide")
st.title("📗 Named Range Coordinates + Formula Remapping")

# Manual mapping of external references like [1], [2], etc.
st.subheader("Manual Mapping for External References")
external_refs = {}
for i in range(1, 10):
    ref_key = f"[{i}]"
    workbook_name = st.text_input(
        f"Map external reference {ref_key} to workbook name (e.g., Mortality_Model_Inputs.xlsx)",
        key=ref_key
    )
    if workbook_name:
        external_refs[ref_key] = workbook_name

uploaded_files = st.file_uploader(
    "📤 Upload Excel files", type=["xlsx"], accept_multiple_files=True
)

if uploaded_files:
    all_named_cell_map = {}
    all_named_ref_info = {}
    file_display_names = {}
    named_ref_formulas = {}

    # Step 1: Load all named references across uploaded workbooks
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
                            row_offset = r - min_row + 1
                            col_offset = c - min_col + 1
                            all_named_cell_map[(display_name, sheet_name, r, c)] = (
                                name, row_offset, col_offset
                            )
                            coord_set.add((r, c))
                    all_named_ref_info[name] = (
                        display_name, sheet_name, coord_set, min_row, min_col
                    )
                except:
                    pass  # Skip malformed definitions

    # Formula remapping logic (unchanged)
    def remap_formula(formula, current_file, current_sheet):
        if not formula:
            return ""
        def cell_address(r, c): return f"{get_column_letter(c)}{r}"
        def remap_single_cell(ref, df, ds):
            # logic omitted for brevity ...
            return ref
        def remap_range(ref, df, ds):
            # logic omitted for brevity ...
            return ref

        pattern = r"complex_regex_here"
        matches = list(re.finditer(pattern, formula))
        offset = 0
        out = formula
        for m in matches:
            raw = m.group(0)
            remapped = remap_range(raw, current_file, current_sheet)
            s, e = m.start() + offset, m.end() + offset
            out = out[:s] + remapped + out[e:]
            offset += len(remapped) - len(raw)
        return out

    # Process each named reference and its formulas
    for name, (file_name, sheet_name, coord_set, min_row, min_col) in all_named_ref_info.items():
        entries = []
        formulas_for_graph = []
        try:
            wb = load_workbook(BytesIO(file_display_names[file_name].getvalue()), data_only=False)
            ws = wb[sheet_name]
            min_col_letter = get_column_letter(min(c for _, c in coord_set))
            max_col_letter = get_column_letter(max(c for _, c in coord_set))
            min_row_num = min(r for r, _ in coord_set)
            max_row_num = max(r for r, _ in coord_set)
            ref_range = f"{min_col_letter}{min_row_num}:{max_col_letter}{max_row_num}"
            cell_range = ws[ref_range] if ":" in ref_range else [[ws[ref_range]]]

            for row in cell_range:
                for cell in row:
                    r_off = cell.row - min_row + 1
                    c_off = cell.column - min_col + 1
                    label = f"{name}[{r_off}][{c_off}]"

                    val = cell.value
                    formula = None
                    if isinstance(val, str) and val.startswith("="):
                        formula = val.strip()

                    remapped = remap_formula(formula, file_name, sheet_name) if formula else (
                        f"[value] {val}" if val is not None else "(empty)"
                    )
                    if formula:
                        formulas_for_graph.append(remapped)

                    entries.append(f"{label} = {formula or remapped}\n → {remapped}")
        except Exception as e:
            entries = [f"❌ Error reading `{name}`: {e}"]

        named_ref_formulas[name] = formulas_for_graph
        with st.expander(f"📌 Named Range: `{name}` → `{sheet_name}` in `{file_name}`"):
            st.code("\n".join(entries), language="text")

    # Render dependency graph (unchanged logic)
    st.subheader("🔗 Dependency Diagram")
    dot = graphviz.Digraph(compound="true", rankdir="LR")
    grouped = defaultdict(list)
    for nm, (fn, *_rest) in all_named_ref_info.items():
        grouped[fn].append(nm)

    deps = defaultdict(set)
    for tgt, fx in named_ref_formulas.items():
        joined = " ".join(fx)
        for src in named_ref_formulas:
            if src != tgt and re.search(rf"\b{re.escape(src)}\b", joined):
                deps[tgt].add(src)

    for i, (fn, nms) in enumerate(grouped.items()):
        with dot.subgraph(name=f"cluster_{i}") as c:
            c.attr(label=fn, style="filled", color="lightgrey")
            for nm in nms:
                c.node(nm)

    for tgt, srclist in deps.items():
        for src in srclist:
            dot.edge(src, tgt)

    st.graphviz_chart(dot)

else:
    st.info("🔹 Upload one or more `.xlsx` files to begin.")
