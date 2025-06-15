import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from io import BytesIO
import re
import os
from collections import defaultdict
import graphviz

st.set_page_config(page_title="Named Range Formula Remapper", layout="wide")
st.title("📔 Named Range Coordinates + Formula Remapping")

# Allow manual mapping of external references like [1], [2], etc.
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

uploaded_files = st.file_uploader("📂 Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_named_cell_map = {}
    all_named_ref_info = {}
    file_display_names = {}
    named_ref_formulas = {}

    # Step 1: Load all named references across uploaded files
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
                            row_off = r - min_row + 1
                            col_off = c - min_col + 1
                            all_named_cell_map[(display_name, sheet_name, r, c)] = (name, row_off, col_off)
                            coord_set.add((r, c))

                    all_named_ref_info[name] = (display_name, sheet_name, coord_set, min_row, min_col)
                except:
                    continue

    # Step 2: Formula remapping logic
    def remap_formula(formula, current_file, current_sheet):
        # Ensure formula is always a string
        if not formula:
            return ""
        formula = str(formula)

        def cell_address(r, c):
            return f"{get_column_letter(c)}{r}"

        def remap_single_cell(ref, default_file, default_sheet):
            if "!" in ref:
                sheet_part, addr = ref.split("!")
                match = re.match(r"\[(\d+)\]", sheet_part)
                if match:
                    ref_key = match.group(0)
                    external_file = external_refs.get(ref_key, ref_key)
                    return f"[{external_file}]{ref}"
                sheet_name = sheet_part
            else:
                sheet_name, addr = default_sheet, ref

            addr = addr.replace("$", "").upper()
            m = re.match(r"([A-Z]+)([0-9]+)", addr)
            if not m:
                return ref

            col, row = column_index_from_string(m.group(1)), int(m.group(2))
            key = (default_file, sheet_name, row, col)
            if key in all_named_cell_map:
                name, roff, coff = all_named_cell_map[key]
                return f"[{default_file}]{name}[{roff}][{coff}]"
            return f"[{default_file}]{sheet_name}!{addr}"

        def remap_range(ref, default_file, default_sheet):
            if ref.startswith("["):
                match = re.match(r"\[(\d+)\]", ref)
                if match:
                    ext = match.group(0)
                    extfile = external_refs.get(ext, ext)
                    return f"[{extfile}]{ref}"

            if "!" in ref:
                sheet, addr = ref.split("!")
            else:
                sheet, addr = default_sheet, ref

            addr = addr.replace("$", "").upper()
            if ":" not in addr:
                return remap_single_cell(ref, default_file, default_sheet)

            start, end = addr.split(":")
            m1, m2 = re.match(r"([A-Z]+)([0-9]+)", start), re.match(r"([A-Z]+)([0-9]+)", end)
            if not (m1 and m2):
                return ref

            sc, sr = column_index_from_string(m1.group(1)), int(m1.group(2))
            ec, er = column_index_from_string(m2.group(1)), int(m2.group(2))
            labels = set()
            for r in range(sr, er + 1):
                for c in range(sc, ec + 1):
                    key = (default_file, sheet, r, c)
                    if key in all_named_cell_map:
                        nm, ro, co = all_named_cell_map[key]
                        labels.add(f"[{default_file}]{nm}[{ro}][{co}]")
                    else:
                        labels.add(f"[{default_file}]{sheet}!{cell_address(r,c)}")
            return ", ".join(sorted(labels))

        pattern = r"(?<![A-Za-z0-9_])(?:\[[^\]]+\])?[A-Za-z0-9_]+!\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Z]{1,3}\$?[0-9]{1,7})?" \
                  r"|(?<![A-Za-z0-9_])\$?[A-Z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Za-z0-9_]{1,3}\$?[0-9]{1,7})?"
        matches = list(re.finditer(pattern, formula))
        new = formula
        offset = 0
        for m in matches:
            raw = m.group(0)
            mapped = remap_range(raw, current_file, current_sheet)
            a, b = m.start() + offset, m.end() + offset
            new = new[:a] + mapped + new[b:]
            offset += len(mapped) - len(raw)

        return new

    # Step 3: Display named references and remapped formulas
    for name, (f, sheet, coord_set, minr, minc) in all_named_ref_info.items():
        entries, formulas_for_graph = [], []
        try:
            wb = load_workbook(BytesIO(file_display_names[f].getvalue()), data_only=False)
            ws = wb[sheet]
            # determine the full range
            minc_letter = get_column_letter(min(c for _, c in coord_set))
            maxc_letter = get_column_letter(max(c for _, c in coord_set))
            minr_num = min(r for r, _ in coord_set)
            maxr_num = max(r for r, _ in coord_set)
            ref_rng = f"{minc_letter}{minr_num}:{maxc_letter}{maxr_num}"
            cr = ws[ref_rng] if ":" in ref_rng else [[ws[ref_rng]]]

            # extract and remap formulas per cell
            for row in cr:
                for cell in row:
                    roff = cell.row - minr + 1
                    coff = cell.column - minc + 1
                    lab = f"{name}[{roff}][{coff}]"
                    frm = None
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        frm = cell.value.strip()
                    elif hasattr(cell.value, "text"):
                        frm = str(cell.value.text).strip()
                    elif cell.value is not None:
                        frm = str(cell.value)
                    rem = remap_formula(frm, f, sheet) if frm else frm
                    formulas_for_graph.append(rem or "")
                    entries.append(f"{lab} = {frm}\n → {rem}")
        except Exception as e:
            entries.append(f"❌ Error reading `{name}` in `{sheet}`: {e}")

        nested = "\n".join(entries)
        st.expander(f"📌 Named Range: `{name}` → sheet `{sheet}` in `{f}`").code(nested)

        named_ref_formulas[name] = formulas_for_graph

    # Step 4: Build and render dependency graph
    st.subheader("🔗 Dependency Graph")
    dot = graphviz.Digraph(compound='true', rankdir='LR')
    grouped = defaultdict(list)
    for nm, (f, *_rest) in all_named_ref_info.items():
        grouped[f].append(nm)

    deps = defaultdict(set)
    for targ, frms in named_ref_formulas.items():
        combined = " ".join(frms)
        for src in named_ref_formulas:
            if src != targ and re.search(rf"\b{re.escape(src)}\b", combined):
                deps[targ].add(src)

    for idx, (fname, nodes) in enumerate(grouped.items()):
        with dot.subgraph(name=f"cluster_{idx}") as c:
            c.attr(label=fname, style='filled', color='lightgrey')
            for nd in nodes:
                c.node(nd)

    for t, ss in deps.items():
        for s in ss:
            dot.edge(s, t)

    st.graphviz_chart(dot)
else:
    st.info("⬆️ Upload one or more `.xlsx` files to begin")
