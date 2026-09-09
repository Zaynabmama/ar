"""Credit AR report generator -- upload the four daily exports, get the
finished 'AR Aging' workbook back.
"""

from __future__ import annotations

import hashlib
import tempfile
from pathlib import Path

import streamlit as st

from ar_report.mailer import (
    load_recipients,
    send_report_email_graph,
    send_report_email_smtp,
)
from ar_report.pipeline import run_pipeline


def _secret(name, default=""):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


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
            workbook_bytes = workbook_buffer.getvalue()
            st.download_button(
                "Download finished workbook",
                data=workbook_buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="credit_ar_download_btn",
            )

            # ── Email the workbook to the fixed distribution list ──────────
            g_tenant = _secret("graph_tenant_id")
            g_client = _secret("graph_client_id")
            g_secret_val = _secret("graph_client_secret")
            g_sender = _secret("graph_sender")
            use_graph = all((g_tenant, g_client, g_secret_val, g_sender))

            recipients = load_recipients()
            if not recipients:
                st.warning(
                    "No recipients configured yet — fill in `ar_report/data/"
                    "recipients.csv` (To, CC, Subject, Body columns)."
                )
            elif use_graph:
                # Guard against re-sending on every Streamlit rerun (e.g.
                # clicking the download button reruns this whole script) -
                # only send once per distinct built workbook.
                sent_key = f"credit_ar_email_sent::{hashlib.md5(workbook_bytes).hexdigest()}"
                if sent_key in st.session_state:
                    st.info(f"📧 Email already sent for this report via **{g_sender}**.")
                    st.table([st.session_state[sent_key]])
                else:
                    with st.spinner(f"Emailing the report to {len(recipients['to'])} recipients..."):
                        try:
                            result = send_report_email_graph(
                                workbook_bytes,
                                file_name,
                                g_tenant,
                                g_client,
                                g_secret_val,
                                g_sender,
                                recipients,
                            )
                            st.session_state[sent_key] = result
                            if result["status"] == "sent":
                                st.success("📧 Sent automatically. ✅")
                            else:
                                st.warning(f"📧 Send failed: {result['status']}")
                            st.table([result])
                        except Exception as mail_err:
                            st.error(f"Sending failed: {mail_err}")
            else:
                with st.expander("📧 Email the report (manual / SMTP fallback)"):
                    st.write("**To:**", ", ".join(a for _, a in recipients["to"]))
                    if recipients["cc"]:
                        st.write("**CC:**", ", ".join(a for _, a in recipients["cc"]))
                    st.write("**Subject:**", recipients["subject"])

                    smtp_host = _secret("smtp_host", "smtp.gmail.com")
                    smtp_port = int(_secret("smtp_port", 465))
                    sender = st.text_input(
                        "Sender email address",
                        value=_secret("gmail_user"),
                        key="credit_ar_mail_user",
                    )
                    app_password = st.text_input(
                        "Email password / App Password",
                        value=_secret("gmail_app_password"),
                        type="password",
                        key="credit_ar_mail_pass",
                    )
                    st.caption(
                        f"Sending via **{smtp_host}:{smtp_port}** (change with "
                        "`smtp_host` / `smtp_port` in Streamlit secrets). Gmail "
                        "needs an App Password: Google Account → Security → "
                        "2-Step Verification → App passwords."
                    )
                    if st.button("Send email now", key="credit_ar_mail_send"):
                        if not (sender and app_password):
                            st.error("Enter the sender address and App Password first.")
                        else:
                            try:
                                with st.spinner("Sending email..."):
                                    result = send_report_email_smtp(
                                        workbook_bytes, file_name, sender,
                                        app_password, recipients,
                                        host=smtp_host, port=smtp_port,
                                    )
                                st.success("Sent. ✅")
                                st.table([result])
                            except Exception as mail_err:
                                st.error(f"Sending failed: {mail_err}")
