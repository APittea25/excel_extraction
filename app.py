import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from io import BytesIO
import re
from collections import defaultdict
import graphviz

st.set_page_config(page_title="Named Range Formula Remapper", layout="wide")
st.title("📘 Named Range Coordinates + Formula Remapping")

# Manual mapping for external workbook references like [1], [2], etc.
st.subheader("Manual Mapping for External References")
external_refs = {}
for i in range(1, 10):
    key = f"[{i}]"
    name = st.text_input(f"Map external reference {key} to workbook name", key=key)
    if name:
        external_refs[key] = name

# Upload Excel files
uploaded_files = st.file_uploader("📁 Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_named_cell_map = {}
    all_named_ref_info = {}
    file_display_names = {}
    named_ref_formulas = {}

    # Build named reference maps
    for uploaded_file in uploaded_files:
        fname = uploaded_file.name
        file_display_names[fname] = uploaded_file
        st.header(f"File: `{fname}`")
        wb = load_workbook(BytesIO(uploaded_file.read()), data_only=False)

        for name in wb.defined_names:
            dn = wb.defined_names[name]
            if dn.is_external or not dn.attr_text:
                continue
            for sheet, ref in dn.destinations:
                ws = wb[sheet]
                ref_clean = ref.replace("$", "").split("!")[-1]
                # Retrieve cells
                cells = ws[ref_clean] if ":" in ref_clean else [[ws[ref_clean]]]
                min_row = min(c.row for r in cells for c in r)
                min_col = min(c.column for r in cells for c in r)

                coords = set()
                for row in cells:
                    for cell in row:
                        r, c = cell.row, cell.column
                        ro = r - min_row + 1
                        co = c - min_col + 1
                        all_named_cell_map[(fname, sheet, r, c)] = (name, ro, co)
                        coords.add((r, c))
                all_named_ref_info[name] = (fname, sheet, coords, min_row, min_col)

    # Formula remapping
    def remap_formula(formula, default_file, default_sheet):
        if not formula:
            return ""

        def cell_addr(r, c):
            return f"{get_column_letter(c)}{r}"

        def remap_single(ref):
            if "!" in ref:
                sheet_p, addr = ref.split("!")
                if sheet_p.startswith("["):
                    match = re.match(r"\[(\d+)\]", sheet_p)
                    if match:
                        key = match.group(0)
                        wbname = external_refs.get(key)
                        return f"[{wbname}]{ref}" if wbname else ref
                sheet_name = sheet_p
            else:
                sheet_name = default_sheet
                addr = ref

            addr = addr.replace("$", "").upper()
            m = re.match(r"([A-Z]+)(\d+)", addr)
            if not m:
                return ref
            col, row = m.group(1), int(m.group(2))
            key = (default_file, sheet_name, row, column_index_from_string(col))
            if key in all_named_cell_map:
                nm, ro, co = all_named_cell_map[key]
                return f"[{default_file}]{nm}[{ro}][{co}]"
            return f"[{default_file}]{sheet_name}!{addr}"

        def remap_range(ref):
            if ref.startswith("["):
                return remap_single(ref)  # external not mapped

            if "!" in ref:
                sheet_name, addr = ref.split("!")
            else:
                sheet_name = default_sheet
                addr = ref

            if ":" not in addr:
                return remap_single(ref)

            start, end = addr.split(":")
            m1, m2 = re.match(r"([A-Z]+)(\d+)", start), re.match(r"([A-Z]+)(\d+)", end)
            if not m1 or not m2:
                return ref

            r1, c1 = int(m1.group(2)), column_index_from_string(m1.group(1))
            r2, c2 = int(m2.group(2)), column_index_from_string(m2.group(1))

            mapped = []
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    key = (default_file, sheet_name, r, c)
                    if key in all_named_cell_map:
                        nm, ro, co = all_named_cell_map[key]
                        mapped.append(f"[{default_file}]{nm}[{ro}][{co}]")
                    else:
                        mapped.append(f"[{default_file}]{sheet_name}!{cell_addr(r, c)}")
            return ", ".join(sorted(set(mapped)))

        pat = r"(?<![A-Za-z0-9_])(?:\[[^\]]+\])?[A-Za-z0-9_]+!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|(?<![A-Za-z0-9_])\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?"
        out = formula
        offset = 0
        for m in re.finditer(pat, formula):
            raw = m.group(0)
            rf = remap_range(raw)
            start, end = m.start() + offset, m.end() + offset
            out = out[:start] + rf + out[end:]
            offset += len(rf) - len(raw)
        return out

    # Collect each named reference's formulas
    for name, (fname, sheet, coords, min_r, min_c) in all_named_ref_info.items():
        entries = []
        graph_frms = []

        wb = load_workbook(BytesIO(file_display_names[fname].getvalue()), data_only=False)
        ws = wb[sheet]
        min_col = min(c for (_, c) in coords)
        max_col = max(c for (_, c) in coords)
        min_row = min(r for (r, _) in coords)
        max_row = max(r for (r, _) in coords)
        rng = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
        cells = ws[rng]

        for row in cells:
            for cell in row:
                ro, co = row.index(cell) + 1, list(cells).index(row) + 1
                lbl = f"{name}[{ro}][{co}]"
                fm = cell.value.strip() if isinstance(cell.value, str) and cell.value.startswith("=") else cell.value or ""
                rm = remap_formula(fm, fname, sheet) if fm else fm
                entries.append(f"{lbl} = {fm}\n→ {rm}")
                if rm:
                    graph_frms.append(rm)

        named_ref_formulas[name] = graph_frms
        with st.expander(f"📌 Named Range: `{name}` → `{sheet}` in `{fname}`"):
            st.code("\n".join(entries), language="text")

    # Fast dependency graph build
    st.subheader("🔗 Dependency Graph")
    dot = graphviz.Digraph(rankdir="LR", compound="true")
    grouped = defaultdict(list)
    for nm, (f, *_r) in all_named_ref_info.items():
        grouped[f].append(nm)
    deps = defaultdict(set)

    for tgt, frms in named_ref_formulas.items():
        all_text = " ".join(frms)
        for src in named_ref_formulas:
            if src != tgt and re.search(rf"\b{re.escape(src)}\b", all_text):
                deps[tgt].add(src)

    for idx, (file, nlist) in enumerate(grouped.items()):
        with dot.subgraph(name=f"cluster_{idx}") as c:
            c.attr(label=file, style="filled", color="lightgrey")
            for n in nlist:
                c.node(n)

    for tgt, sl in deps.items():
        for s in sl:
            dot.edge(s, tgt)

    st.graphviz_chart(dot)

else:
    st.info("⬆️ Upload Excel files (.xlsx) to begin.")
