import io

import streamlit as st

from sales_report.logic import build, filter_quarter, load_data, summary, write


@st.cache_data(show_spinner=False)
def _load_cached(data_bytes: bytes):
    return load_data(io.BytesIO(data_bytes))


@st.cache_data(show_spinner=False)
def _build_cached(data_bytes: bytes, year: int | None, quarter: int | None):
    lines = _load_cached(data_bytes)
    if year and quarter:
        lines = filter_quarter(lines, year, quarter)
    result = build(lines)
    out = io.BytesIO()
    write(result, out)
    return out.getvalue(), summary(result), dict(result["new_terms"])


def _quarters(lines):
    seen = set()
    for r in lines:
        d = r.get("Doc Dt")
        if hasattr(d, "year"):
            seen.add((d.year, (d.month - 1) // 3 + 1))
    return sorted(seen, reverse=True)


def render_sales_by_pt_tool():
    st.markdown("### Sales by Payment Terms")
    upload = st.file_uploader(
        "Upload **Sales margin report**", type=["xlsx", "xlsm"], key="sales_pt_uploader"
    )
    if not upload:
        return

    data_bytes = upload.getvalue()
    try:
        with st.spinner("Reading the report..."):
            lines = _load_cached(data_bytes)
    except Exception as exc:  # noqa: BLE001 - show the reason to the user
        st.error(f"Could not read the file: {exc}")
        return

    quarters = _quarters(lines)
    options = [f"{y} Q{q}" for y, q in quarters] + ["All rows (no date filter)"]
    choice = st.selectbox("Quarter (by Doc Dt)", options, key="sales_pt_quarter")
    if choice.startswith("All"):
        year = quarter = None
    else:
        year, quarter = int(choice[:4]), int(choice[-1])

    with st.spinner("Building the report..."):
        out_bytes, s, new_terms = _build_cached(data_bytes, year, quarter)

    n_inv, total = s["All invoices"]
    st.success(f"{n_inv:,} invoices, total Adjusted Sales **{total:,.2f}**.")
    st.table(
        [
            {"Bucket": k, "Invoices": f"{v[0]:,}", "Total": f"{v[1]:,.2f}"}
            for k, v in s.items()
            if isinstance(v, tuple)
        ]
    )
    if new_terms:
        st.warning(
            f"{len(new_terms)} pay term(s) are not in the rules list. They were put "
            "in **Open** and are listed on the 'New pay terms' sheet. Add them to "
            "`sales_report/data/pay_terms.csv` with the right bucket."
        )
        st.table([{"Pay term": t, "Invoices": n, "Amount": f"{a:,.2f}"} for t, (n, a) in new_terms.items()])

    label = choice.replace(" ", "-") if year else "all"
    st.download_button(
        label="Download Sales by Payment Terms.xlsx",
        data=out_bytes,
        file_name=f"Sales by Payment Terms {label}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="sales_pt_download",
    )
