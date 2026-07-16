import streamlit as st

from BUM.logic import build_bum_workbook


@st.cache_data(show_spinner=False)
def _build_cached(
    main_bytes: bytes,
    auh_bytes: bytes,
    renewals_bytes: bytes,
    insurance_bytes: bytes,
    pdc_bytes: bytes,
    backlog_bytes: bytes,
):
    return build_bum_workbook(
        main_bytes,
        auh_bytes,
        renewals_bytes,
        insurance_bytes,
        pdc_bytes,
        backlog_bytes,
    )


def render_bum_tool():
    st.markdown("### BUM Report Builder")

    st.caption(
        "Upload the **main AR dump** (as-of date in A1, headers on row 2) plus the "
        "five reports. The tool adds the aging and lookup columns "
        "(Invoice Age ... PDC) as live Excel formulas, computes the PDC and "
        "Backlog pivots, and bundles the fixed BUM / Region / GSI / SE Africa "
        "lists into the output file."
    )

    col1, col2 = st.columns(2)
    with col1:
        main_upload = st.file_uploader(
            "Upload **Main AR file**",
            type=["xlsx", "xlsm"],
            key="bum_main_uploader",
        )
        auh_upload = st.file_uploader(
            "Upload **AUH customers list**",
            type=["xlsx", "xlsm"],
            key="bum_auh_uploader",
        )
        renewals_upload = st.file_uploader(
            "Upload **Renewals invoices**",
            type=["xlsx", "xlsm"],
            key="bum_renewals_uploader",
        )
    with col2:
        insurance_upload = st.file_uploader(
            "Upload **Insurance master**",
            type=["xlsx", "xlsm"],
            key="bum_insurance_uploader",
        )
        pdc_upload = st.file_uploader(
            "Upload **PDC due to be banked**",
            type=["xlsx", "xlsm"],
            key="bum_pdc_uploader",
        )
        backlog_upload = st.file_uploader(
            "Upload **Sales Backlog report**",
            type=["xlsx", "xlsm"],
            key="bum_backlog_uploader",
        )

    uploads = (
        main_upload,
        auh_upload,
        renewals_upload,
        insurance_upload,
        pdc_upload,
        backlog_upload,
    )
    if not all(uploads):
        return

    try:
        with st.spinner("Building BUM master file..."):
            out_bytes, meta = _build_cached(*(u.getvalue() for u in uploads))

        st.success(
            f"As-of **{meta['as_of']:%d-%b-%Y}** (end of month "
            f"{meta['eom']:%d-%b-%Y}) — {meta['rows']:,} invoice rows, "
            f"{meta['auh_rows']:,} AUH customers, "
            f"{meta['renewal_rows']:,} renewal invoices, "
            f"{meta['insurance_rows']:,} insured customers."
        )
        st.caption(
            f"PDC pivot: {meta['pdc_customers']:,} customers, "
            f"total {meta['pdc_total']:,.2f}. "
            f"Backlog pivot: {meta['backlog_customers']:,} customers, "
            f"total {meta['backlog_total']:,.2f}."
        )
        st.caption(
            "Check cell: on Sheet1 row 1, the value under **9. AR Balance** "
            "must be ~0 (new AR Balance vs original Ar Balance)."
        )

        st.download_button(
            label="Download BUM Master.xlsx",
            data=out_bytes,
            file_name=f"BUM Master {meta['as_of']:%Y-%m-%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="bum_download",
        )

    except Exception as e:
        st.error(
            f"{e}\n\nIf this persists, expand 'Details' for traceback and share the top 10 lines."
        )
        st.exception(e)
