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
        "Upload one or more **By_Customer** files (Orion or Traverse output) and the Insurance Master. "
        "The output is the QBR quarterly provision forecast model (ALL sheet) "
        "with live formulas; quarters ending on or before the AR Data Date are inactive. "
        "When multiple By_Customer files are uploaded, their rows are combined as-is into a single "
        "output (no merging of duplicate customers across files), and all files must detect the "
        "same starting quarter."
    )

    bud_uploads = st.file_uploader(
        "Upload **By_Customer** Excel (one or more files)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
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

    if not bud_uploads:
        return

    try:
        with st.spinner("Reading By_Customer file(s)..."):
            loaded = [
                (f.name, *_read_by_customer_workbook(f.getvalue())) for f in bud_uploads
            ]

        quarters_detected = {name: q for name, _, _, q in loaded}
        distinct_quarters = set(quarters_detected.values())
        if len(distinct_quarters) > 1:
            mismatch_lines = "\n".join(f"- **{name}**: {q}" for name, q in quarters_detected.items())
            st.error(
                "Uploaded files disagree on the detected starting quarter - they must all be "
                f"from the same quarter to combine into one output:\n\n{mismatch_lines}"
            )
            return

        selected_quarter = loaded[0][3]

        for name, _, sheet_name, _ in loaded:
            st.success(f"Loaded **{name}** — sheet: {sheet_name}")

        st.caption(
            f"Detected starting quarter: **{selected_quarter}** "
            "(used to pre-fill the Collections FC columns; earlier quarters stay blank)"
        )

        # ── Compute mapped rows per file, then combine as-is (shared between Export and Dashboard) ──
        with st.spinner("Mapping data..."):
            mapped_parts = []
            fallback_files = []
            for name, df_customer_only, _, _ in loaded:
                part = _map_bud_rows_cached(
                    df_customer_only,
                    ins_df=master_df,
                    selected_quarter=selected_quarter,
                )
                if not part.attrs.get("used_not_due_breakdown", True):
                    fallback_files.append(name)
                mapped_parts.append(part)

            bud_rows = pd.concat(mapped_parts, ignore_index=True)
            bud_rows.attrs["used_not_due_breakdown"] = not fallback_files

        if fallback_files:
            st.warning(
                "The following file(s) have no 'Not Due 0-30 / 31-60 / ...' columns - their "
                "whole Not Due amount was placed in 'Not Due 0-30 days': "
                f"{', '.join(fallback_files)}. Regenerate them with tool 1 for a proper breakdown."
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
