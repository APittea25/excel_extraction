import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from io import BytesIO
import re
import os
from collections import defaultdict
import graphviz

st.set_page_config(page_title="Named Range Formula Remapper", layout="wide")
st.title("📘 Named Range Coordinates + Formula Remapping")

# Auto-expand all expanders on load
st.markdown("""
    <script>
    window.addEventListener('load', function() {
        document.querySelectorAll('details').forEach(el => el.open = true);
    });
    </script>
""", unsafe_allow_html=True)

# Manual mapping for external workbook refs
st.subheader("Manual Mapping for External References")
external_refs = {}
for i in range(1, 10):
    key = f"[{i}]"
    name = st.text_input(f"Map external reference {key} → workbook name (e.g., MyWorkbook.xlsx)", key=key)
    if name:
        external_refs[key] = name

uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_named_cell_map = {}
    all_named_ref_info = {}
    file_contents = {}
    named_ref_formulas = {}

    for uploaded in uploaded_files:
        fname = uploaded.name
        file_contents[fname] = uploaded.getvalue()
        st.header(f"File: `{fname}`")
        wb = load_workbook(BytesIO(file_contents[fname]), data_only=False)
        for name in wb.defined_names:
            dn = wb.defined_names[name]
            if dn.is_external or not dn.attr_text:
                continue
            for sheet, ref in dn.destinations:
                try:
                    ws = wb[sheet]
                    addr = ref.replace("$", "").split("!")[-1]
                    cells = ws[addr] if ":" in addr else [[ws[addr]]]
                    min_r = min(c.row for row in cells for c in row)
                    min_c = min(c.column for row in cells for c in row)
                    coords = set()
                    for row in cells:
                        for cell in row:
                            key = (fname, sheet, cell.row, cell.column)
                            row_off = cell.row - min_r + 1
                            col_off = cell.column - min_c + 1
                            all_named_cell_map[key] = (name, row_off, col_off)
                            coords.add((cell.row, cell.column))
                    all_named_ref_info[name] = (fname, sheet, coords, min_r, min_c)
                except:
                    pass

    def remap_formula(formula, ffile, fsheet):
        if not formula:
            return ""
        def cell_addr(r, c): return f"{get_column_letter(c)}{r}"
        
        def remap_one(ref):
            if "!" in ref:
                sheetp, addr = ref.split("!")
                match = re.match(r"\[(\d+)\]$", sheetp)
                if match:
                    ext = match.group(0)
                    extname = external_refs.get(ext, ext)
                    return f"[{extname}]{ref}"
                sheet = sheetp
            else:
                sheet, addr = fsheet, ref

            addr = addr.replace("$", "").upper()
            m = re.match(r"([A-Z]+)([0-9]+)", addr)
            if not m:
                return ref
            r, c = int(m[2]), column_index_from_string(m[1])
            key = (ffile, sheet, r, c)
            if key in all_named_cell_map:
                nm, ro, co = all_named_cell_map[key]
                return f"[{ffile}]{nm}[{ro}][{co}]"
            return f"[{ffile}]{sheet}!{addr}"

        def remap_range(ref):
            if ref.startswith("["):
                m = re.match(r"\[(\d+)\]", ref)
                if m:
                    ext = m.group(0)
                    extname = external_refs.get(ext, ext)
                    return f"[{extname}]{ref}"

            if ":" not in ref:
                return remap_one(ref)
            start, end = ref.split(":")
            kv = set(remap_one(start).split(", ") + remap_one(end).split(", "))
            return ", ".join(sorted(kv))

        pat = r"(?<![A-Za-z0-9_])(?:\[[^\]]+\])?[A-Za-z0-9_]+!\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?|(?<![A-Za-z0-9_])\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Za-z0-9_]{1,3}\$?[0-9]{1,7})?"
        new = ""
        last = 0
        for m in re.finditer(pat, formula):
            new += formula[last:m.start()] + remap_range(m.group(0))
            last = m.end()
        new += formula[last:]
        return new

    for name, (ffile, fsheet, coords, mr, mc) in all_named_ref_info.items():
        entries = []
        formulas = []
        wb = load_workbook(BytesIO(file_contents[ffile]), data_only=False)
        ws = wb[fsheet]
        min_c = min(c for (_, c) in coords)
        max_c = max(c for (_, c) in coords)
        min_r = min(r for (r, _) in coords)
        max_r = max(r for (r, _) in coords)
        ridx = f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c)}{max_r}"

        for row in ws[ridx]:
            for cell in row:
                r_off = cell.row - mr + 1
                c_off = cell.column - mc + 1
                label = f"{name}[{r_off}][{c_off}]"
                try:
                    val = cell.value
                    formula = val.strip() if isinstance(val, str) and val.startswith("=") else f"[value] {val}"
                    rem = remap_formula(formula, ffile, fsheet)
                    if formula.startswith("="):
                        formulas.append(rem)
                except Exception as e:
                    rem = f"[error: {e}]"
                entries.append(f"{label} = {formula}\n → {rem}")
        named_ref_formulas[name] = formulas
        with st.expander(f"{name} → {fsheet} in {ffile}"):
            st.code("\n".join(entries), language="text")

    # Draw dependency graph
    st.subheader("🔗 Dependency Graph")
    dot = graphviz.Digraph(graph_attr={"compound": "true", "rankdir": "LR"})
    groups = defaultdict(list)
    for nm, (ffile, *_a) in all_named_ref_info.items():
        groups[ffile].append(nm)

    deps = defaultdict(set)
    for tgt, forms in named_ref_formulas.items():
        txt = " ".join(forms)
        for src in named_ref_formulas:
            if src != tgt and re.search(rf"\b{re.escape(src)}\b", txt):
                deps[tgt].add(src)

    for idx, (ff, nms) in enumerate(groups.items()):
        with dot.subgraph(name=f"cluster_{idx}") as c:
            c.attr(label=ff, style="filled", color="lightgrey")
            for nm in nms:
                c.node(nm)

    for tgt, sources in deps.items():
        for src in sources:
            dot.edge(src, tgt)

    st.graphviz_chart(dot)

else:
    st.info("Upload one or more `.xlsx` files to begin.")
