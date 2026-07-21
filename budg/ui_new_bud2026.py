import datetime as dt

import pandas as pd
import streamlit as st

from budg.bud2026_dashboard import render_dashboard
from budg.bud2026_export import export_bud2026_quarterly
from budg.bud2026_mapper import map_by_customer_to_bud2026
from budg.insurance_master import load_insurance_master
from common.identifier_utils import normalize_excel_identifier_series
from common.quarter_utils import detect_selected_quarter_from_columns


@st.cache_data(show_spinner=False)
def _read_by_customer_workbook(file_bytes: bytes):
    xl = pd.ExcelFile(pd.io.common.BytesIO(file_bytes), engine="openpyxl")
    sheet_name = "By_Customer" if "By_Customer" in xl.sheet_names else xl.sheet_names[0]
    df_customer_only = pd.read_excel(xl, sheet_name=sheet_name)
    if "Main Ac" in df_customer_only.columns:
        df_customer_only["Main Ac"] = normalize_excel_identifier_series(df_customer_only["Main Ac"])
    selected_quarter = detect_selected_quarter_from_columns(df_customer_only.columns)
    return df_customer_only, sheet_name, selected_quarter


@st.cache_data(show_spinner=False)
def _load_insurance_master_cached(file_bytes: bytes):
    return load_insurance_master(pd.io.common.BytesIO(file_bytes))


@st.cache_data(show_spinner=False)
def _map_bud_rows_cached(df_customer_only: pd.DataFrame, ins_df: pd.DataFrame | None, selected_quarter: str):
    return map_by_customer_to_bud2026(
        df_customer_only,
        ins_df=ins_df,
        selected_quarter=selected_quarter,
    )


def render_new_bud_tool():
    st.markdown("### BUD2026 Builder (Quarterly Model)")

    st.caption(
        "Upload the **By_Customer** and the Insurance Master. "
        "The output is the QBR quarterly provision forecast model (ALL sheet) "
        "with live formulas; quarters ending on or before the AR Data Date are inactive."
    )

    bud_upload = st.file_uploader(
        "Upload **By_Customer** Excel",
        type=["xlsx", "xls"],
        key="new_uploader",
    )

    ins_upload = st.file_uploader(
        "Upload **Insurance Master** Excel",
        type=["xlsx", "xls"],
        key="ins_uploader",
    )

    ar_date = st.date_input(
        "AR Data Date",
        value=dt.date.today(),
        min_value=dt.date(2025, 1, 1),
        max_value=dt.date(2027, 12, 31),
        format="DD/MM/YYYY",
        key="bud_ar_date",
        help="Written to ALL!B5. Quarters ending on or before this date return 0 "
        "(their collections/provisions are already inside the AR data).",
    )

    master_df = None
    if ins_upload:
        with st.spinner("Loading Insurance Master..."):
            master_df = _load_insurance_master_cached(ins_upload.getvalue())

        st.success(
            f"Insurance Master loaded: {len(master_df)} unique (Customer Code, Main Account)"
        )

    if not bud_upload:
        return

    try:
        with st.spinner("Reading By_Customer..."):
            df_customer_only, sheet_name, selected_quarter = _read_by_customer_workbook(
                bud_upload.getvalue()
            )

        st.success(f"Loaded sheet: {sheet_name}")

        st.caption(
            f"Detected starting quarter from By_Customer columns: **{selected_quarter}** "
            "(used to pre-fill the Collections FC columns; earlier quarters stay blank)"
        )

        # ── Compute mapped rows (shared between Export and Dashboard) ──────
        with st.spinner("Mapping data..."):
            bud_rows = _map_bud_rows_cached(
                df_customer_only,
                ins_df=master_df,
                selected_quarter=selected_quarter,
            )

        if not bud_rows.attrs.get("used_not_due_breakdown", True):
            st.warning(
                "The By_Customer sheet has no 'Not Due 0-30 / 31-60 / ...' columns - "
                "the whole Not Due amount was placed in 'Not Due 0-30 days'. "
                "Regenerate the By_Customer file with tool 1 for a proper breakdown."
            )

        st.markdown("---")

        # ── Tabs: Export  |  Management Dashboard ─────────────────────────
        tab_export, tab_dashboard = st.tabs(["📥 Export", "📊Dashboard with AI analysis"])

        with tab_export:
            st.subheader("Export BUD2026 Quarterly Model")

            with st.spinner("Building workbook..."):
                bud_bytes = export_bud2026_quarterly(bud_rows, ar_date)

            st.download_button(
                label="Download AR Collection and Provision Forecast - Quarterly.xlsx",
                data=bud_bytes,
                file_name="AR Collection and Provision Forecast - Quarterly.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="new_single_download",
            )

        with tab_dashboard:
            render_dashboard(bud_rows, selected_quarter=selected_quarter)

    except Exception as e:
        st.error(
            f"{e}\n\nIf this persists, expand 'Details' for traceback and share the top 10 lines."
        )
        st.exception(e)
