import pandas as pd
import streamlit as st

from budg.bud2026_export import export_bud2026_ordered
from budg.bud2026_headers import (
    BANNER_ANCHORS_BUD2026,
    HEADERS_BUD2026,
    filter_banner_anchors_by_headers,
    filter_headers_by_quarter,
)
from budg.bud2026_mapper import map_by_customer_to_bud2026
from budg.insurance_master import load_insurance_master
from common.quarter_utils import detect_selected_quarter_from_columns


def render_new_bud_tool():
    st.markdown("### BUD2026 Builder")

    st.caption(
        "Upload ONLY the **By_Customer** sheet. "
        "Optionally upload an **Insurance Master**; the loader will auto-detect Budg-style "
        "or Traverse-style columns."
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

    master_df = None
    if ins_upload:
        with st.spinner("Loading Insurance Master..."):
            master_df = load_insurance_master(ins_upload)

        st.success(
            f"Insurance Master loaded: {len(master_df)} unique (Customer Code, Main Account)"
        )
        st.dataframe(master_df.head(10))

    if not bud_upload:
        return

    try:
        with st.spinner("Reading By_Customer..."):
            xl = pd.ExcelFile(bud_upload, engine="openpyxl")
            sheet_name = "By_Customer" if "By_Customer" in xl.sheet_names else xl.sheet_names[0]
            df_customer_only = pd.read_excel(xl, sheet_name=sheet_name)

        st.success(f"Loaded sheet: {sheet_name}")
        st.dataframe(df_customer_only.head(20), use_container_width=True)

        st.markdown("---")
        st.subheader("Export BUD2026")

        selected_quarter = detect_selected_quarter_from_columns(df_customer_only.columns)
        st.caption(
            f"Detected starting quarter from By_Customer columns: **{selected_quarter}**"
        )

        bud_rows = map_by_customer_to_bud2026(df_customer_only, ins_df=master_df)
        export_headers = filter_headers_by_quarter(HEADERS_BUD2026.copy(), selected_quarter)
        export_banners = filter_banner_anchors_by_headers(BANNER_ANCHORS_BUD2026, export_headers)

        bud_bytes = export_bud2026_ordered(
            bud_rows,
            export_headers,
            banner_anchors=export_banners,
            header_gap_rows=1,
            freeze=True,
            autofilter=True,
            merge_banner=True,
        )

        st.download_button(
            label="Download AR Collection and Provision Forecast BUD2026.xlsx",
            data=bud_bytes,
            file_name="AR Collection and Provision Forecast BUD2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="new_single_download",
        )

    except Exception as e:
        st.error(
            f"{e}\n\nIf this persists, expand 'Details' for traceback and share the top 10 lines."
        )
        st.exception(e)
