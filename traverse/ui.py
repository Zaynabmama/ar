from datetime import date

import streamlit as st

from common.quarter_utils import QUARTER_ORDER
from traverse.export import export_traverse_ar
from traverse.processor import prepare_traverse_input, prepare_traverse_output


def get_as_of_date_for_quarter(selected_quarter: str) -> date:
    """Return locked as_of_date based on quarter selection."""
    quarter_dates = {
        "Q1": date(2026, 1, 1),   # Start of year
        "Q2": date(2026, 3, 31),  # End of Q1
        "Q3": date(2026, 6, 30),  # End of Q2
        "Q4": date(2026, 9, 30),  # End of Q3
    }
    return quarter_dates.get(selected_quarter, date(2026, 1, 1))


def render_traverse_tool():
    st.header("Traverse -> AR Forecast Tool")

    selected_quarter = st.selectbox(
        "Starting Quarter",
        options=QUARTER_ORDER,
        index=0,
        key="traverse_quarter_selection",
        help="Kept for the same flow as Orion. We can reuse it when the downstream sheets are added.",
    )

    as_of_date = get_as_of_date_for_quarter(selected_quarter)
    st.caption(f"As of Date (auto-locked): **{as_of_date.strftime('%d/%m/%Y')}**")

    source_currency = st.radio(
        "Source Currency",
        options=["JD", "USD"],
        horizontal=True,
        index=0,
        key="traverse_source_currency",
        help="Choose JD if GrossAmount still needs conversion to USD. Choose USD if GrossAmount is already in USD.",
    )

    st.caption(f"Selected quarter: {selected_quarter}")

    uploaded_file = st.file_uploader(
        "Upload Traverse Excel",
        type=["xlsx", "xls"],
        key="traverse_uploader",
    )

    if not uploaded_file:
        return

    try:
        with st.spinner("Processing Traverse file..."):
            df_raw = prepare_traverse_input(uploaded_file)
            df_out = prepare_traverse_output(df_raw)

        st.success(f"Loaded Traverse input with {len(df_out)} rows.")
        preview_df = df_out.loc[:, ~df_out.columns.duplicated(keep="last")]
        st.dataframe(preview_df.head(20), use_container_width=True)

        st.subheader("Download")
        excel_file = export_traverse_ar(
            df_out,
            as_of_date=as_of_date,
            selected_quarter=selected_quarter,
            source_currency=source_currency,
        )

        st.download_button(
            "Download Traverse_AR.xlsx",
            data=excel_file.getvalue(),
            file_name="Traverse_AR.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="traverse_download_btn",
        )

    except Exception as e:
        st.error("An error occurred while processing the Traverse file.")
        st.exception(e)
