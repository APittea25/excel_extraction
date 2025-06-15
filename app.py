import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from io import BytesIO
import re
from collections import defaultdict
import graphviz

st.set_page_config(page_title="Named Range Formula Remapper", layout="wide")
st.title("📘 Named Range Coordinates + Formula Remapping")

# Expand all sections on load
st.markdown("""
<script>
window.addEventListener('load', () => document.querySelectorAll('details').forEach(el => el.open = true));
</script>
""", unsafe_allow_html=True)

# Manual mapping UI for external references [1]–[9]
st.subheader("🛠️ Manual Mapping for External References")
external_refs = {}
for i in range(1, 10):
    key = f"[{i}]"
    workbook_name = st.text_input(f"{key} → workbook name", key=key)
    if workbook_name:
        external_refs[key] = workbook_name

uploaded_files = st.file_uploader("📂 Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_named_cell_map = {}
    all_named_ref_info = {}
    file_uploads = {}
    named_ref_formulas = {}

    # Step 1: Load named references from each workbook
    for uploaded_file in uploaded_files:
        fname = uploaded_file.name
        file_uploads[fname] = uploaded_file
        st.header(f"📄 File: `{fname}`")
        wb = load_workbook(BytesIO(uploaded_file.read()), data_only=False)

        for nm in wb.defined_names:
            dn = wb.defined_names[nm]
            if dn.is_external or not dn.attr_text:
                continue
            for sht, ref in dn.destinations:
                try:
                    ws = wb[sht]
                    rc = ref.replace("$", "").split("!")[-1]
                    cells = ws[rc] if ":" in rc else [[ws[rc]]]
                    min_r = min(c.row for row in cells for c in row)
                    min_c = min(c.column for row in cells for c in row)
                    coords = set()
                    for row in cells:
                        for c in row:
                            key = (fname, sht, c.row, c.column)
                            all_named_cell_map[key] = (
                                nm,
                                c.row - min_r + 1,
                                c.column - min_c + 1
                            )
                            coords.add((c.row, c.column))
                    all_named_ref_info[nm] = (fname, sht, coords, min_r, min_c)
                except:
                    continue

    # Step 2: Formula remapping logic
    def remap_formula(formula, curr_file, curr_sheet):
        if not formula:
            return ""
        def addr(r, c): return f"{get_column_letter(c)}{r}"

        def adjust_external(raw):
            if raw.startswith("[") and "]!" in raw:
                idx = raw.split("]")[0] + "]"
                return raw.replace(idx, f"[{external_refs.get(idx, idx)}]")
            return raw

        pattern = re.compile(
            r"(?:\[[^\]]+\]!)?[A-Za-z0-9_]+!\$?[A-Z]{1,3}\$?[0-9]+"
            r"(?::\$?[A-Z]{1,3}\$?[0-9]+)?"
        )
        out = formula
        shift = 0

        for m in pattern.finditer(formula):
            raw = adjust_external(m.group(0))
            parts = raw.split("!")
            sheet = curr_sheet if len(parts) == 1 else parts[-2].strip("[]")
            cell_ref = parts[-1].replace("$", "")
            if ":" in cell_ref:
                start, _ = cell_ref.split(":")
                col, row = re.match(r"([A-Z]+)([0-9]+)", start).groups()
                key = (curr_file, sheet, int(row), column_index_from_string(col))
                if key in all_named_cell_map:
                    nm, ro, co = all_named_cell_map[key]
                    raw = f"[{curr_file}]{nm}[{ro}][{co}]"
            else:
                col, row = re.match(r"([A-Z]+)([0-9]+)", cell_ref).groups()
                key = (curr_file, sheet, int(row), column_index_from_string(col))
                if key in all_named_cell_map:
                    nm, ro, co = all_named_cell_map[key]
                    raw = f"[{curr_file}]{nm}[{ro}][{co}]"

            s, e = m.start() + shift, m.end() + shift
            out = out[:s] + raw + out[e:]
            shift += len(raw) - (e - s)

        return out

    # Step 3: Extract and remap formulas per named reference
    for nm, (f, sht, coords, min_r, min_c) in all_named_ref_info.items():
        wb = load_workbook(BytesIO(file_uploads[f].getvalue()), data_only=False)
        ws = wb[sht]
        entries = []
        formulas = []

        for r, c in sorted(coords):
            cell = ws[f"{get_column_letter(c)}{r}"]
            val = cell.value
            raw_formula = val if isinstance(val, str) and val.startswith("=") else str(val)
            remap = remap_formula(raw_formula, f, sht)
            entries.append(f"{nm}[{r-min_r+1}][{c-min_c+1}] = {raw_formula}\n → {remap}")
            if remap:
                formulas.append(remap)

        named_ref_formulas[nm] = formulas
        with st.expander(f"📌 `{nm}` → `{sht}` in `{f}`"):
            st.code("\n".join(entries), language="text")

    # Step 4: Stable & faster Dependency Graph
    st.subheader("🔗 Dependency Graph")
    dot = graphviz.Digraph()
    dot.attr(rankdir="LR")

    groups = defaultdict(list)
    for nm, (f, sht, *_rest) in all_named_ref_info.items():
        groups[f].append(nm)

    for idx, (f, nms) in enumerate(groups.items()):
        sub = graphviz.Digraph(f"cluster_{idx}", graph_attr={
            'label': f, 'style': 'filled', 'color': 'lightgrey'
        })
        for nm in nms:
            sub.node(nm, color='blue')
        dot.subgraph(sub)

    all_names = set(named_ref_formulas)
    for tgt, fs in named_ref_formulas.items():
        txt = " ".join(fs)
        for src in all_names - {tgt}:
            if re.search(rf"\b{re.escape(src)}\b", txt):
                dot.edge(src, tgt)

    st.graphviz_chart(dot)

else:
    st.info("⬆️ Upload Excel files to begin processing.")
