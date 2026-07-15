import time
from datetime import date, timedelta

import pandas as pd
import streamlit as st

from budg.insurance_master import load_insurance_master
from provision.export import export_provision_forecast
from provision.mapper import map_by_customer_to_provision


@st.cache_data(show_spinner=False)
def _read_tool1_workbook(file_bytes: bytes):
    xl = pd.ExcelFile(pd.io.common.BytesIO(file_bytes), engine="openpyxl")
    by_customer_sheet = "By_Customer" if "By_Customer" in xl.sheet_names else xl.sheet_names[0]
    df_customer = pd.read_excel(xl, sheet_name=by_customer_sheet)
    return df_customer, by_customer_sheet


@st.cache_data(show_spinner=False)
def _load_insurance_master_cached(file_bytes: bytes):
    return load_insurance_master(pd.io.common.BytesIO(file_bytes))


@st.cache_data(show_spinner=False)
def _build_workbook_cached(file_bytes: bytes, ins_bytes: bytes | None, ar_date: date):
    df_customer, by_customer_sheet = _read_tool1_workbook(file_bytes)
    ins_df = _load_insurance_master_cached(ins_bytes) if ins_bytes else None
    df_fixed, used_breakdown = map_by_customer_to_provision(df_customer, ins_df=ins_df)
    workbook_bytes = export_provision_forecast(df_fixed, ar_date)
    return workbook_bytes, by_customer_sheet, len(df_fixed), used_breakdown


def _default_ar_date() -> date:
    today = date.today()
    return today.replace(day=1) - timedelta(days=1)


def render_provision_tool():
    st.markdown("### AR Collection & Provision Forecast")
    st.caption(
        "Upload the output of the **AR Backlog** tool (the By_Customer workbook), "
        "set the **AR Data Date**, and download the provision forecast model with live formulas."
    )

    ar_date = st.date_input(
        "AR Data Date",
        value=_default_ar_date(),
        key="prov_ar_date",
        help="Master control (cell B3). Months on or before this date return 0; "
        "the first month after it becomes the active month.",
    )

    main_upload = st.file_uploader(
        "Upload **By_Customer** workbook (output of the AR Backlog tool)",
        type=["xlsx", "xls"],
        key="prov_uploader",
    )

    ins_upload = st.file_uploader(
        "Upload **Insurance Master** Excel (optional - fills the Insurance column)",
        type=["xlsx", "xls"],
        key="prov_ins_uploader",
    )

    if ins_upload:
        try:
            master_df = _load_insurance_master_cached(ins_upload.getvalue())
            st.success(f"Insurance Master loaded: {len(master_df)} unique (Customer Code, Main Account)")
        except Exception as e:
            st.error(f"Insurance Master could not be read: {e}")
            return

    if not main_upload:
        return

    if ar_date.year != 2026:
        st.warning(
            f"You selected **{ar_date:%d-%b-%Y}**, but this model is built for **FY 2026** "
            "(all monthly columns are Jan-Dec 2026). A date outside 2026 will produce an "
            "empty or meaningless forecast."
        )

    try:
        total_start = time.perf_counter()
        with st.spinner("Building provision forecast workbook..."):
            workbook_bytes, by_customer_sheet, n_customers, used_breakdown = _build_workbook_cached(
                main_upload.getvalue(),
                ins_upload.getvalue() if ins_upload else None,
                ar_date,
            )
        elapsed = time.perf_counter() - total_start

        st.success(f"Loaded sheet: {by_customer_sheet} - {n_customers} customers.")

        if used_breakdown:
            st.info(
                "Not Due breakdown (0-30 / 31-60 / 61-90 / 91-180 / 180+) read from the "
                "By_Customer 'Not Due ...' columns (collectible view)."
            )
        else:
            st.warning(
                "The By_Customer sheet has no 'Not Due 0-30 / 31-60 / ...' columns - the whole "
                "Not Due amount was placed in 'Not Due 0-30 days' (column K). Regenerate the "
                "file with the latest AR Backlog tool to get the breakdown."
            )

        st.download_button(
            "Download AR Collection and Provision Forecast",
            data=workbook_bytes,
            file_name=f"AR Collection and Provision Forecast - {ar_date:%B}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="prov_download_btn",
        )
        st.caption(f"Generated in {elapsed:.1f} sec. Formulas recalculate when the file opens in Excel.")

    except Exception as e:
        st.error("An error occurred while building the provision forecast.")
        st.exception(e)
