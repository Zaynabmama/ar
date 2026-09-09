import streamlit as st

from ar_report.ui import render_ar_report_tool
from BUM.ui import render_bum_tool
from budg.ui_new_bud2026 import render_new_bud_tool
from orion.ui import render_orion_tool
from orion_monthly.ui import render_orion_monthly_tool
from provision.ui import render_provision_tool
from traverse.ui import render_traverse_tool


st.set_page_config(page_title="AR Backlogr", layout="wide")
st.title("AR Backlog")

tab_old, tab_monthly, tab_new, tab_provision, tab_bum, tab_credit_ar = st.tabs(
    ["AR Backlog", "AR Backlog (Monthly)", "BUD2026 from By_Customer", "AR Provision Forecast", "BUM", "Credit AR Report"]
)

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

with tab_monthly:
    render_orion_monthly_tool()

with tab_new:
    render_new_bud_tool()

with tab_provision:
    render_provision_tool()

with tab_bum:
    render_bum_tool()

with tab_credit_ar:
    render_ar_report_tool()
