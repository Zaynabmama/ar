import streamlit as st

from budg.ui_new_bud2026 import render_new_bud_tool
from orion.ui import render_orion_tool
from traverse.ui import render_traverse_tool


st.set_page_config(page_title="AR Backlog Ultra-Fast Processor", layout="wide")
st.title("AR Backlog")

tab_old, tab_new = st.tabs(["AR Backlog", "BUD2026 from By_Customer"])

with tab_old:
    source = st.radio(
        "Choose source",
        options=["Orion", "Traverse"],
        horizontal=True,
        key="ar_source_selector",
    )

    if source == "Orion":
        render_orion_tool()
    else:
        render_traverse_tool()

with tab_new:
    render_new_bud_tool()
