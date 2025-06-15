import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from io import BytesIO
import re
from collections import defaultdict
import graphviz

# —– Session state for expand/collapse toggle —–
if "expanded_all" not in st.session_state:
    st.session_state.expanded_all = False

def toggle():
    st.session_state.expanded_all = not st.session_state.expanded_all

st.set_page_config(page_title="Named Range Formula Remapper", layout="wide")
st.title("📘 Named Range Formula Remapper")
st.button("🔁 Expand / Collapse All Named Ranges", on_click=toggle)

# —– Manual mapping UI —–
st.subheader("Manual Mapping for External References")
external_refs = {}
for i in range(1, 10):
    key = f"[{i}]"
    result = st.text_input(f"Map external {key} → workbook name (e.g., Mortality_Model_Inputs.xlsx)", key=key)
    if result:
        external_refs[key] = result

uploaded_files = st.file_uploader("Upload .xlsx files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_named_cell_map = {}
    all_named_ref_info = {}
    file_display_names = {}
    named_ref_formulas = {}

    # —– Load files, extract named ranges —–
    for uploaded_file in uploaded_files:
        fname = uploaded_file.name
        file_display_names[fname] = uploaded_file
        st.header(f"📄 File: `{fname}`")
        wb = load_workbook(BytesIO(uploaded_file.read()), data_only=False)

        for name in wb.defined_names:
            dn = wb.defined_names[name]
            if dn.is_external or not dn.attr_text:
                continue
            for sheet, ref in dn.destinations:
                try:
                    ws = wb[sheet]
                    ref_clean = ref.replace("$", "").split("!")[-1]
                    cells = ws[ref_clean] if ":" in ref_clean else [[ws[ref_clean]]]
                    min_row = min(cell.row for row in cells for cell in row)
                    min_col = min(cell.column for row in cells for cell in row)

                    coord_set = set()
                    for row in cells:
                        for cell in row:
                            r, c = cell.row, cell.column
                            row_off = r - min_row + 1
                            col_off = c - min_col + 1
                            key = (fname, sheet, r, c)
                            all_named_cell_map[key] = (name, row_off, col_off)
                            coord_set.add((r, c))

                    all_named_ref_info[name] = (fname, sheet, coord_set, min_row, min_col)
                except:
                    continue

    # —– Helper for remapping formulas —–
    def remap_formula(formula, current_file, current_sheet):
        if not formula:
            return ""

        def cell_address(r, c):
            return f"{get_column_letter(c)}{r}"

        def remap_cell(ref, dfile, dsheet):
            if "!" in ref:
                sheet_part, addr = ref.split("!")
                m = re.match(r"\[(\d+)\]", sheet_part)
                if m:
                    ext = m.group(0)
                    ext_file = external_refs.get(ext, ext)
                    return f"[{ext_file}]{ref}"
                sheet = sheet_part
            else:
                sheet = dsheet
                addr = ref

            addr = addr.replace("$", "").upper()
            m = re.match(r"([A-Z]+)(\d+)", addr)
            if not m:
                return ref
            col_str, row_str = m.groups()
            row = int(row_str)
            col = column_index_from_string(col_str)
            key = (dfile, sheet, row, col)

            if key in all_named_cell_map:
                nm, ro, co = all_named_cell_map[key]
                return f"[{dfile}]{nm}[{ro}][{co}]"
            else:
                return f"[{dfile}]{sheet}!{addr}"

        def remap_range(ref, dfile, dsheet):
            if ref.startswith("["):
                m = re.match(r"\[(\d+)\]", ref)
                if m:
                    return f"[{external_refs.get(m.group(0), m.group(0))}]{ref}"
            if "!" in ref:
                sheet, addr = ref.split("!")
            else:
                sheet, addr = dsheet, ref

            addr = addr.replace("$", "").upper()
            if ":" not in addr:
                return remap_cell(ref, dfile, dsheet)

            start, end = addr.split(":")
            m1 = re.match(r"([A-Z]+)(\d+)", start)
            m2 = re.match(r"([A-Z]+)(\d+)", end)
            if not (m1 and m2):
                return ref

            sc, sr = column_index_from_string(m1.group(1)), int(m1.group(2))
            ec, er = column_index_from_string(m2.group(1)), int(m2.group(2))
            label_set = set()
            for r in range(sr, er+1):
                for c in range(sc, ec+1):
                    key = (dfile, sheet, r, c)
                    if key in all_named_cell_map:
                        nm, ro, co = all_named_cell_map[key]
                        label_set.add(f"[{dfile}]{nm}[{ro}][{co}]")
                    else:
                        label_set.add(f"[{dfile}]{sheet}!{cell_address(r,c)}")
            return ", ".join(sorted(label_set))

        pat = r"(?<![A-Za-z0-9_])(?:\[[^\]]+\])?[A-Za-z0-9_]+!\$?[A-Z]{1,3}\$?[0-9]+(?::\$?[A-Z]{1,3}\$?[0-9]+)?"
        out, offset = formula, 0
        for m in re.finditer(pat, formula):
            raw = m.group(0)
            rem = remap_range(raw, current_file, current_sheet)
            s, e = m.start()+offset, m.end()+offset
            out = out[:s] + rem + out[e:]
            offset += len(rem) - len(raw)

        return out

    # —– Extract formulas & display expanders —–
    for nm, (fname, sheet, coords, minr, minc) in all_named_ref_info.items():
        entries, formulas_for_graph = [], []
        try:
            wb = load_workbook(BytesIO(file_display_names[fname].getvalue()), data_only=False)
            ws = wb[sheet]
            minc_letter = get_column_letter(min(c for _,c in coords))
            maxc_letter = get_column_letter(max(c for _,c in coords))
            minr_num = min(r for r,_ in coords)
            maxr_num = max(r for r,_ in coords)
            rng = f"{minc_letter}{minr_num}:{maxc_letter}{maxr_num}"
            cell_rows = ws[rng] if ":" in rng else [[ws[rng]]]

            for row in cell_rows:
                for cell in row:
                    roff, coff = cell.row - minr + 1, cell.column - minc + 1
                    label = f"{nm}[{roff}][{coff}]"
                    val, remapped = None, None

                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        val = cell.value.strip()
                    elif hasattr(cell.value, "text"):
                        val = str(cell.value.text).strip()
                    elif cell.value is not None:
                        val = "[value] " + str(cell.value)

                    remapped = remap_formula(val, fname, sheet) if val else ""
                    formulas_for_graph.append(remapped)
                    entries.append(f"{label} = {val}\n→ {remapped}")

        except Exception as e:
            entries.append(f"❌ Error processing `{nm}`: {e}")

        named_ref_formulas[nm] = formulas_for_graph
        with st.expander(f"📌 {nm} → {sheet} @ {fname}", expanded=st.session_state.expanded_all):
            st.code("\n".join(entries), language="text")

    # —– Dependency graph —–
    st.subheader("🔗 Dependency Graph")
    dot = graphviz.Digraph()
    dot.attr(compound="true", rankdir="LR")

    grouped = defaultdict(list)
    for nm, (fname, *_rest) in all_named_ref_info.items():
        grouped[fname].append(nm)

    deps = defaultdict(set)
    for tgt, formulas in named_ref_formulas.items():
        text = " ".join(formulas)
        for src in named_ref_formulas:
            if src != tgt and re.search(rf"\b{re.escape(src)}\b", text):
                deps[tgt].add(src)

    for i, (fname, nodes) in enumerate(grouped.items()):
        with dot.subgraph(name=f"cluster_{i}") as c:
            c.attr(label=fname, style="filled", color="lightgrey")
            for nm in nodes:
                c.node(nm)

    for tgt, sources in deps.items():
        for src in sources:
            dot.edge(src, tgt)

    st.graphviz_chart(dot)

else:
    st.info("Upload one or more `.xlsx` files to begin.")
