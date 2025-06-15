import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from io import BytesIO
import re
import graphviz
from collections import defaultdict

st.set_page_config(page_title="Named Range Formula Remapper", layout="wide")
st.title("\U0001F4D8 Named Range Coordinates + Formula Remapping")

uploaded = st.file_uploader("Upload Excel files", type="xlsx", accept_multiple_files=True)

# Initialize flag
if "expanded_all" not in st.session_state:
    st.session_state.expanded_all = False

# Show button only after uploading
if uploaded:
    if st.button("🔁 Expand/Collapse All Named Ranges"):
        st.session_state.expanded_all = not st.session_state.expanded_all

# Inject JS *after* upload/button so it applies to the existing DOM
expanded_flag = "true" if st.session_state.expanded_all else "false"
st.markdown(f"""
    <script>
      const setAll = () => {{
        document.querySelectorAll('details').forEach(el => el.open = {expanded_flag});
      }};
      // Delay to ensure Streamlit DOM is built
      setTimeout(setAll, 100);
    </script>
""", unsafe_allow_html=True)

# Manual mapping inputs
st.subheader("Manual Mapping for External References")
external_refs = {}
for i in range(1, 10):
    key = f"[{i}]"
    val = st.text_input(f"Map external ref {key}", key=key)
    if val:
        external_refs[key] = val

if uploaded:
    # ======== Workbook parsing & remapping (your existing logic) =========
    all_named_cell_map = {}
    all_named_ref_info = {}
    file_display_names = {}
    named_ref_formulas = {}
    # ... load workbooks, remap_formulas, and store named_ref_formulas ...
    # ======================================================================

    # Display named ranges
    for name, (file_name, sheet_name, coord_set, min_row, min_col) in all_named_ref_info.items():
        entries = []
        try:
            # ... reading cells and remapping logic ...
            pass
        except Exception as e:
            entries.append(f"❌ Error accessing `{name}`: {e}")
        with st.expander(f"📌 {name} → {sheet_name} ({file_name})", expanded=st.session_state.expanded_all):
            st.code("\n".join(entries), language="text")

    # Dependency graph
    st.subheader("🔗 Dependency Graph")
    dot = graphviz.Digraph()
    dot.attr(compound='true', rankdir='LR')
    # your grouping and edge addition logic...
    st.graphviz_chart(dot)
else:
    st.info("⬆️ Upload `.xlsx` files to begin")
