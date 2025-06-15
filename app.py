import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from io import BytesIO
import re
from collections import defaultdict
import graphviz

st.set_page_config(page_title="Named Range Formula Remapper", layout="wide")
st.title("📄 Named Range Coordinates + Formula Remapping")

uploaded_files = st.file_uploader("📂 Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_named_cell_map = {}
    all_named_ref_info = {}
    file_display_names = {}
    named_ref_formulas = {}

    # Build reference mappings
    for uploaded_file in uploaded_files:
        fname = uploaded_file.name
        file_display_names[fname] = uploaded_file
        wb = load_workbook(BytesIO(uploaded_file.read()), data_only=False)

        for name in wb.defined_names:
            dn = wb.defined_names[name]
            if dn.is_external or not dn.attr_text:
                continue
            for sheet, ref in dn.destinations:
                try:
                    ws = wb[sheet]
                    clean_ref = ref.replace("$", "").split("!")[-1]
                    cells = ws[clean_ref] if ":" in clean_ref else [[ws[clean_ref]]]

                    min_row = min(cell.row for row in cells for cell in row)
                    min_col = min(cell.column for row in cells for cell in row)
                    coords = set()

                    for row in cells:
                        for cell in row:
                            r, c = cell.row, cell.column
                            r_off = r - min_row + 1
                            c_off = c - min_col + 1
                            all_named_cell_map[(fname, sheet, r, c)] = (name, r_off, c_off)
                            coords.add((r, c))

                    all_named_ref_info[name] = (fname, sheet, coords, min_row, min_col)

                except:
                    continue

    # Remap logic
    def remap_formula(formula, curr_file, curr_sheet):
        if not formula:
            return ""
        def addr(r, c): return f"{get_column_letter(c)}{r}"

        def remap_range(ref):
            sheet = curr_sheet
            addr_str = ref
            if "!" in ref:
                sheet, addr_str = ref.split("!")
            start, end = (addr_str.replace("$", "").split(":") + [None])[:2]

            def single(a):
                col, row = re.match(r"([A-Z]+)([0-9]+)", a).groups()
                r, c = int(row), column_index_from_string(col)
                key = (curr_file, sheet, r, c)
                if key in all_named_cell_map:
                    name, ro, co = all_named_cell_map[key]
                    return f"{name}[{ro}][{co}]"
                return f"{sheet}!{addr(a)}"

            if end:
                s1 = single(start)
                s2 = single(end)
                return f"{s1}:{s2}"
            return single(start)

        pattern = r"(?<=^|[^A-Za-z0-9_])([A-Za-z0-9_!:\$\[\]]+)"
        tokens = re.findall(pattern, formula)
        for t in tokens:
            if re.match(r'\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?', t):
                formula = formula.replace(t, remap_range(t))
        return formula

    # Extract formulas
    for name, (fname, sheet, coords, min_r, min_c) in all_named_ref_info.items():
        formulas_list = []
        wb = load_workbook(BytesIO(file_display_names[fname].getvalue()), data_only=False)
        ws = wb[sheet]
        for (r, c) in coords:
            cell = ws[f"{get_column_letter(c)}{r}"]
            val = None
            if isinstance(cell.value, str) and cell.value.startswith("="):
                val = cell.value
            elif hasattr(cell, 'value'):
                val = str(cell.value)
            mapped = remap_formula(val or "", fname, sheet)
            formulas_list.append(mapped)
        named_ref_formulas[name] = formulas_list

        # Display cells and mapping
        st.subheader(f"🔖 Named Range: {name}  (in {sheet} @ {fname})")
        for i, fstr in enumerate(formulas_list, start=1):
            st.text(f"• {name}[{i}] → {fstr}")

    # Optimized Dependency Graph
    token_index = defaultdict(set)
    for tgt, forms in named_ref_formulas.items():
        for f in forms:
            tokens = re.findall(r'\b[A-Za-z0-9_]+\b', f)
            for t in set(tokens):
                if t in named_ref_formulas:
                    token_index[t].add(tgt)

    dependencies = defaultdict(set)
    for name, tgts in token_index.items():
        for t1 in tgts:
            for t2 in tgts:
                if t1 != t2:
                    dependencies[t1].add(t2)

    # Render graph
    st.subheader("🌐 Dependency Graph")
    dot = graphviz.Digraph()
    dot.attr(compound='true', rankdir='LR')
    groups = defaultdict(list)
    for nm, (fname, *_rest) in all_named_ref_info.items():
        groups[fname].append(nm)

    for idx, (fname, names) in enumerate(groups.items()):
        with dot.subgraph(name=f"cluster_{idx}") as c:
            c.attr(label=fname, style="filled", color="lightgrey")
            for nm in names:
                dot.node(nm)

    for src, tgts in dependencies.items():
        for tgt in tgts:
            dot.edge(src, tgt)
    st.graphviz_chart(dot)

else:
    st.info("⬆️ Upload Excel files to process.")
