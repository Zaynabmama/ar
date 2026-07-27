import streamlit as st

from BUM.logic import build_bum_workbook
from BUM.split import build_bum_zip


@st.cache_data(show_spinner=False)
def _build_zip_cached(*input_bytes):
    return build_bum_zip(*input_bytes)


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
        "Upload the **main AR dump** (banner rows are fine — the tool finds the "
        "'Cust Code' header row and the 'As on Date' automatically; internal "
        "customers (Mindware/Aklaniat/iFix), main accounts 12302/12304/12306 "
        "and the summary row are excluded) plus the five reports. The tool adds "
        "the aging and lookup columns (Invoice Age ... PDC) as live Excel "
        "formulas, computes the PDC and Backlog pivots, and bundles the fixed "
        "BUM / Region / GSI / SE Africa lists into the output file."
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

        with st.spinner("Building per-BUM files..."):
            zip_bytes, zmeta = _build_zip_cached(*(u.getvalue() for u in uploads))

        all_ok = all(c["match"] for c in zmeta["checks"])
        if all_ok:
            st.success(
                "All files verified: the total AR **By Customer** matches the "
                "total AR **By Invoice** in every file. ✅"
            )
        else:
            st.error("Totals do NOT match in some files — check the table below!")

        st.table(
            [
                {
                    "File": c["file"],
                    "Invoices": c["invoices"],
                    "Customers": c["customers"],
                    "Total AR (By Invoice)": f"{c['inv_total']:,.2f}",
                    "Total AR (By Customer)": f"{c['cus_total']:,.2f}",
                    "Matching": "✅" if c["match"] else "❌",
                }
                for c in zmeta["checks"]
            ]
        )
        st.download_button(
            label="Download BUM files (ZIP)",
            data=zip_bytes,
            file_name=f"BUM Files {zmeta['as_of']:%Y-%m-%d}.zip",
            mime="application/zip",
            key="bum_zip_download",
        )

        # ── Email the files to each BUM ────────────────────────────────
        import hashlib

        from BUM.mailer import load_email_map, send_bum_emails, send_bum_emails_graph

        def _gsecret(name, default=""):
            try:
                return st.secrets.get(name, default)
            except Exception:
                return default

        g_tenant = _gsecret("graph_tenant_id")
        g_client = _gsecret("graph_client_id")
        g_secret_val = _gsecret("graph_client_secret")
        g_sender = _gsecret("graph_sender")
        use_graph = all((g_tenant, g_client, g_secret_val, g_sender))

        if use_graph:
            email_map = load_email_map()
            if not email_map:
                st.warning(
                    "No recipients configured yet — fill in `BUM/data/"
                    "emails.csv` (File, To, CC, Body columns)."
                )
            else:
                # Guard against re-sending the same batch on every Streamlit
                # rerun (e.g. clicking another download button reruns this
                # whole script) - only send once per distinct built ZIP.
                sent_key = f"bum_emails_sent::{hashlib.md5(zip_bytes).hexdigest()}"
                if sent_key in st.session_state:
                    st.info(f"📧 Emails already sent for this batch via **{g_sender}**.")
                    st.table(st.session_state[sent_key])
                else:
                    with st.spinner(
                        f"Emailing all {len(email_map)} files via Microsoft 365..."
                    ):
                        try:
                            results = send_bum_emails_graph(
                                zip_bytes,
                                zmeta["as_of"],
                                g_tenant,
                                g_client,
                                g_secret_val,
                                g_sender,
                                email_map,
                            )
                            st.session_state[sent_key] = results
                            ok = sum(1 for r in results if r["status"] == "sent")
                            if ok == len(results):
                                st.success(f"📧 Sent all {ok} emails automatically. ✅")
                            else:
                                st.warning(
                                    f"📧 Sent {ok} of {len(results)} emails — "
                                    "see the table below."
                                )
                            st.table(results)
                        except Exception as mail_err:
                            st.error(f"Sending failed: {mail_err}")
            return

        with st.expander("📧 Email the files to each BUM (manual / SMTP fallback)"):
            email_map = load_email_map()
            if not email_map:
                st.warning(
                    "No recipients configured yet — fill in the email column "
                    "of `BUM/data/emails.csv` (File label → email address)."
                )
            else:
                st.table(
                    [
                        {
                            "File": label,
                            "To": ", ".join(a for _, a in info["to"]),
                            "CC": ", ".join(a for _, a in info["cc"]) or "—",
                        }
                        for label, info in email_map.items()
                    ]
                )
                def _secret(name, default=""):
                    try:
                        return st.secrets.get(name, default)
                    except Exception:
                        return default

                smtp_host = _secret("smtp_host", "smtp.gmail.com")
                smtp_port = int(_secret("smtp_port", 465))
                sender = st.text_input(
                    "Sender email address",
                    value=_secret("gmail_user"),
                    key="bum_mail_user",
                )
                app_password = st.text_input(
                    "Email password / App Password",
                    value=_secret("gmail_app_password"),
                    type="password",
                    key="bum_mail_pass",
                )
                st.caption(
                    f"Sending via **{smtp_host}:{smtp_port}** (change with "
                    "`smtp_host` / `smtp_port` in Streamlit secrets). Gmail "
                    "needs an App Password: Google Account → Security → "
                    "2-Step Verification → App passwords."
                )
                if st.button(
                    f"Send {len(email_map)} emails now", key="bum_mail_send"
                ):
                    if not (sender and app_password):
                        st.error("Enter the Gmail address and App Password first.")
                    else:
                        try:
                            with st.spinner("Sending emails..."):
                                results = send_bum_emails(
                                    zip_bytes, zmeta["as_of"], sender,
                                    app_password, email_map,
                                    host=smtp_host, port=smtp_port,
                                )
                            st.success(f"Sent {len(results)} emails. ✅")
                            st.table(results)
                        except Exception as mail_err:
                            st.error(
                                f"Sending failed: {mail_err}\n\n"
                                "Most common cause: using the normal Gmail "
                                "password instead of an App Password, or "
                                "2-Step Verification not enabled."
                            )

    except Exception as e:
        st.error(
            f"{e}\n\nIf this persists, expand 'Details' for traceback and share the top 10 lines."
        )
        st.exception(e)
