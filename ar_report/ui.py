"""Credit AR report generator -- upload the four daily exports, get the
finished 'AR Aging' workbook back.
"""

from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from ar_report.pipeline import run_pipeline


def render_ar_report_tool():
    st.header("Credit AR Report Generator")
    

    col1, col2 = st.columns(2)
    with col1:
        ar_file = st.file_uploader("AR Provision report (from ORION)", type=["xlsx"], key="credit_ar_ar_uploader")
        insurance_file = st.file_uploader("Insurance report (from ORION)", type=["xlsx"], key="credit_ar_insurance_uploader")
    with col2:
        pdc_file = st.file_uploader("PDC file (from ORION)", type=["xlsx"], key="credit_ar_pdc_uploader")
        backlog_file = st.file_uploader("Sales Backlog report (from email)", type=["xlsx"], key="credit_ar_backlog_uploader")

    region_override_files = st.file_uploader(
        "Region override lists ",
        type=["xlsx"],
        accept_multiple_files=True,
        key="credit_ar_region_override_uploader",
        help="Any customer listed in one of these files gets that file's region "
        "(e.g. AUH) instead of the country-based default. Upload one file per "
        "city/sub-region; leave empty to use the country as-is.",
    )

    as_on_date_override = st.date_input(
        "As on date (leave as-is to use the date inside the AR file)",
        value=None,
        key="credit_ar_as_on_date",
    )

    generate = st.button(
        "Generate Report",
        type="primary",
        disabled=not all([ar_file, pdc_file, insurance_file, backlog_file]),
        key="credit_ar_generate_btn",
    )

    if not generate:
        return

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        paths = {}
        for name, upload in [
            ("ar", ar_file),
            ("pdc", pdc_file),
            ("insurance", insurance_file),
            ("backlog", backlog_file),
        ]:
            p = tmp_path / f"{name}.xlsx"
            p.write_bytes(upload.getbuffer())
            paths[name] = p

        region_override_paths = []
        for i, upload in enumerate(region_override_files or []):
            p = tmp_path / f"region_override_{i}.xlsx"
            p.write_bytes(upload.getbuffer())
            region_override_paths.append(p)

        try:
            with st.spinner("Cleaning data, rebuilding pivots, and assembling the workbook..."):
                workbook_buffer, stats = run_pipeline(
                    ar_path=paths["ar"],
                    pdc_path=paths["pdc"],
                    insurance_path=paths["insurance"],
                    backlog_path=paths["backlog"],
                    as_on_date=as_on_date_override,
                    region_override_paths=region_override_paths,
                )
        except Exception as exc:
            st.error(f"Couldn't generate the report: {exc}")
        else:
            st.success(f"Report generated for {stats['as_on_date'].date()}")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("AR rows kept", f"{stats['ar_rows_clean']:,}", delta=f"-{stats['ar_rows_raw'] - stats['ar_rows_clean']:,} filtered")
            c2.metric("Backlog rows kept", f"{stats['backlog_rows_clean']:,}", delta=f"-{stats['backlog_rows_raw'] - stats['backlog_rows_clean']:,} filtered")
            c3.metric("Customers", f"{stats['customers']:,}")
            c4.metric("Region overrides applied", f"{stats['region_overrides_applied']:,}")

            file_name = f"AR Aging as on {stats['as_on_date'].strftime('%d.%m.%Y')}.xlsx"
            st.download_button(
                "Download finished workbook",
                data=workbook_buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="credit_ar_download_btn",
            )
