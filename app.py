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

# Manual mapping UI
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

    # Step 1: Load named references
    for uploaded_file in uploaded_files:
        f = uploaded_file.name
        file_uploads[f] = uploaded_file
        st.header(f"📄 File: `{f}`")
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
                            key = (f, sht, c.row, c.column)
                            all_named_cell_map[key] = (nm, c.row - min_r + 1, c.column - min_c + 1)
                            coords.add((c.row, c.column))
                    all_named_ref_info[nm] = (f, sht, coords, min_r, min_c)
                except:
                    continue

    # Step 2: Remapping
    def remap_formula(formula, curr_file, curr_sheet):
        if not formula: return ""
        def cell_address(r, c): return f"{get_column_letter(c)}{r}"

        def adjust_external(raw):
            if raw.startswith("[") and "]!" in raw:
                idx = raw.split("]")[0] + "]"
                if idx in external_refs:
                    return raw.replace(idx, f"[{external_refs[idx]}]")
            return raw

        pattern = re.compile(r"(?:\[[^\]]+\]!)?[A-Za-z0-9_]+!\$?[A-Z]{1,3}\$?[0-9]+(?::\$?[A-Z]{1,3}\$?[0-9]+)?")
        modified = formula
        shift = 0

        for m in pattern.finditer(formula):
            raw = m.group(0)
            raw = adjust_external(raw)

            parts = raw.split("!")
            sheet = curr_sheet if len(parts) == 1 else parts[-2].strip("[]")
            addr = parts[-1].replace("$", "")
            if ":" in addr:
                start, end = addr.split(":")
                for cell in (start, end):
                    col, row = re.match(r"([A-Z]+)([0-9]+)", cell).groups()
                    key = (curr_file, sheet, int(row), column_index_from_string(col))
                    if key in all_named_cell_map:
                        nm, ro, co = all_named_cell_map[key]
                        raw = f"[{curr_file}]{nm}[{ro}][{co}]"
                        break
            else:
                col, row = re.match(r"([A-Z]+)([0-9]+)", addr).groups()
                key = (curr_file, sheet, int(row), column_index_from_string(col))
                if key in all_named_cell_map:
                    nm, ro, co = all_named_cell_map[key]
                    raw = f"[{curr_file}]{nm}[{ro}][{co}]"

            start, end = m.start() + shift, m.end() + shift
            modified = modified[:start] + raw + modified[end:]
            shift += len(raw) - (end - start)

        return modified

    # Step 3: Extract remapped formulas
    for nm, (f, sht, coords, min_r, min_c) in all_named_ref_info.items():
        try:
            wb = load_workbook(BytesIO(file_uploads[f].getvalue()), data_only=False)
            ws = wb[sht]
            formulas = []
            entries = []

            for r, c in sorted(coords):
                cell = ws[f"{get_column_letter(c)}{r}"]
                v = cell.value if isinstance(cell.value, str) and cell.value.startswith("=") else None
                rem = remap_formula(v or str(cell.value), f, sht)
                entries.append(f"{nm}[{r-min_r+1}][{c-min_c+1}] = {v}\n → {rem}")
                if rem: formulas.append(rem)

            named_ref_formulas[nm] = formulas
            with st.expander(f"📌 `{nm}` → `{sht}` in `{f}`"):
                st.code("\n".join(entries), language="text")
        except Exception as e:
            st.error(f"Error in {nm}: {e}")

    # Step 4: Build optimized dependency graph
    st.subheader("🔗 Dependency Graph")
    dot = graphviz.Digraph(rankdir="LR")

    groups = defaultdict(list)
    for nm, (f, sht, *_rest) in all_named_ref_info.items():
        groups[f].append((nm, sht))

    for idx, (f, nodes) in enumerate(groups.items()):
        with dot.subgraph(f"cluster_{idx}") as c:
            c.attr(label=f, style="filled", color="lightgrey")
            for nm, sht in nodes:
                col = "blue"  # Optionally color by sheet
                c.node(nm, color=col)

    all_names = set(named_ref_formulas)
    for tgt, forms in named_ref_formulas.items():
        text = " ".join(forms)
        for src in all_names:
            if src != tgt and re.search(rf"\b{re.escape(src)}\b", text):
                dot.edge(src, tgt)

    st.graphviz_chart(dot)

else:
    st.info("⬆️ Upload Excel files to begin processing.")
