import time

import streamlit as st

from common.month_utils import MONTH_ORDER_2026, MONTH_SELECT_LABELS
from orion.export import fast_excel_download_monthly
from orion.processor import customer_summary_monthly, invoice_summary, process_ar_file


def render_orion_monthly_tool():
    st.header("AR Backlog -> By_Customer Forecast Tool (Monthly)")

    selected_month = st.selectbox(
        "Starting Month",
        options=MONTH_ORDER_2026,
        format_func=lambda m: MONTH_SELECT_LABELS[m],
        index=0,
        key="monthly_tool_month_selection",
        help="Choose the month the uploaded AR Backlog should start forecasting from.",
    )

    uploaded_file = st.file_uploader(
        "Upload AR Backlog Excel",
        type=["xlsx", "xls"],
        key="monthly_uploader",
        help="Upload the AR Backlog workbook (As on Date in B14, header row = 16).",
    )

    if not uploaded_file:
        return

    try:
        total_start = time.perf_counter()

        process_start = time.perf_counter()
        with st.spinner("Processing file..."):
            df_main = process_ar_file(uploaded_file)
            df_customer = customer_summary_monthly(df_main, selected_month=selected_month)
            df_invoice = invoice_summary(df_main)

        process_end = time.perf_counter()
        st.success("Processing completed.")

        st.subheader("Download")
        excel_file = fast_excel_download_monthly(
            df_main,
            df_customer,
            df_invoice,
            selected_month=selected_month,
        )

        st.download_button(
            "Download Processed File",
            data=excel_file.getvalue(),
            file_name="processed_AR_backlog_monthly.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="monthly_download_btn",
        )

        st.subheader("Performance Metrics")
        c1, c2, c3 = st.columns(3)
        c1.metric("Processing Time", f"{process_end - process_start:.2f} sec")
        export_end = time.perf_counter()
        c2.metric("Export Time", f"{export_end - process_end:.2f} sec")
        total_end = time.perf_counter()
        c3.metric("Total Runtime", f"{total_end - total_start:.2f} sec")

    except Exception as e:
        st.error("An error occurred.")
        st.exception(e)
