"""
budg/bud2026_dashboard.py

Management dashboard for BUD2026.
- AI-powered risk flag detector
- AR balance by customer status (donut)
- Top exposures by AR balance (table)
"""

import json

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from budg.azure_openai import (
    get_azure_openai_config,
    render_azure_openai_settings,
    run_azure_openai_text,
)

_AR_BAL = " AR\nBalance"
_INSURANCE = "Insurance"
_A121_150 = "Aging\n121 to 150"
_AGE_GE_151 = "Aging\n>=151"
_PROV_INCL_INS = "Provision after collection including Insurance/BG/LC"
_OPEN_PROV = "AR Provision at\n31/08/2025"
_REGION = "Sales Budget region"
_STATUS = "Customer Status"
_CUST_NAME = "Cust Name"
_FOCUS = "Focus List"
_FC_Q1 = "Collections FC\n31/03/2026"

_CHART_LAYOUT = dict(
    margin=dict(l=0, r=0, t=8, b=0),
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(size=12),
)


def _num(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series(0.0, index=df.index)
    return pd.to_numeric(df[col], errors="coerce").fillna(0.0)


def _current_collection_header(selected_quarter: str) -> str:
    mapping = {
        "Q1": "Collections FC\n31/03/2026",
        "Q2": "Collections FC\n30/06/2026",
        "Q3": "Collections FC\n30/09/2026",
        "Q4": "Collections FC\n31/12/2026",
    }
    return mapping.get(selected_quarter, _FC_Q1)


def _build_data_summary(df: pd.DataFrame, selected_quarter: str) -> dict:
    total_ar = float(_num(df, _AR_BAL).sum())
    total_prov = float(_num(df, _PROV_INCL_INS).sum())
    open_prov = float(_num(df, _OPEN_PROV).sum())
    prov_diff = total_prov - open_prov
    current_fc_header = _current_collection_header(selected_quarter)
    current_fc = float(_num(df, current_fc_header).sum())
    coverage = round(total_prov / total_ar * 100, 1) if total_ar else 0
    late_ar = float((_num(df, _A121_150) + _num(df, _AGE_GE_151)).sum())
    late_pct = round(late_ar / total_ar * 100, 1) if total_ar else 0

    status_split = {}
    if _STATUS in df.columns:
        for status, group in df.groupby(_STATUS):
            status_split[str(status)] = round(float(_num(group, _AR_BAL).sum()) / 1_000_000, 2)

    region_split = {}
    if _REGION in df.columns:
        for region, group in df.groupby(_REGION):
            region_split[str(region)] = round(float(_num(group, _AR_BAL).sum()) / 1_000_000, 2)

    working = df.copy()
    working["__ar"] = _num(working, _AR_BAL)
    working["__ins"] = _num(working, _INSURANCE)
    working["__late"] = _num(working, _A121_150) + _num(working, _AGE_GE_151)
    working["__current_fc"] = _num(working, current_fc_header)
    working["__prov"] = _num(working, _PROV_INCL_INS)

    top10 = []
    for _, row in working.nlargest(10, "__ar").iterrows():
        top10.append(
            {
                "name": str(row.get(_CUST_NAME, "")),
                "region": str(row.get(_REGION, "")),
                "status": str(row.get(_STATUS, "")),
                "focus_list": str(row.get(_FOCUS, "")),
                "ar_m": round(float(row["__ar"]) / 1_000_000, 2),
                "insurance_m": round(float(row["__ins"]) / 1_000_000, 2),
                "late_ar_m": round(float(row["__late"]) / 1_000_000, 2),
                "current_fc_m": round(float(row["__current_fc"]) / 1_000_000, 2),
                "provision_m": round(float(row["__prov"]) / 1_000_000, 2),
            }
        )

    uninsured = working[(working["__late"] > 0) & (working["__ins"] == 0)].nlargest(5, "__late")
    uninsured_list = [
        {
            "name": str(row.get(_CUST_NAME, "")),
            "late_ar_m": round(float(row["__late"]) / 1_000_000, 2),
            "status": str(row.get(_STATUS, "")),
            "region": str(row.get(_REGION, "")),
        }
        for _, row in uninsured.iterrows()
    ]

    focus_list = []
    if _FOCUS in df.columns:
        focus_df = working[
            working[_FOCUS].astype(str).str.strip().str.upper().isin(["YES", "Y", "X", "1", "TRUE"])
        ]
        for _, row in focus_df.iterrows():
            focus_list.append(
                {
                    "name": str(row.get(_CUST_NAME, "")),
                    "ar_m": round(float(row["__ar"]) / 1_000_000, 2),
                    "current_fc_m": round(float(row["__current_fc"]) / 1_000_000, 2),
                    "status": str(row.get(_STATUS, "")),
                }
            )

    return {
        "selected_quarter": selected_quarter,
        "current_fc_label": current_fc_header,
        "total_ar_m": round(total_ar / 1_000_000, 2),
        "total_provision_m": round(total_prov / 1_000_000, 2),
        "opening_provision_m": round(open_prov / 1_000_000, 2),
        "provision_change_m": round(prov_diff / 1_000_000, 2),
        "current_fc_collections_m": round(current_fc / 1_000_000, 2),
        "coverage_pct": coverage,
        "late_ar_pct": late_pct,
        "late_ar_m": round(late_ar / 1_000_000, 2),
        "customer_count": int(len(df)),
        "status_breakdown_m": status_split,
        "region_breakdown_m": region_split,
        "top_10_customers": top10,
        "uninsured_late_ar": uninsured_list,
        "focus_list_customers": focus_list,
    }


def _ai_chat_completion(messages: list[dict], max_tokens: int = 800, temperature: float = 0.2) -> str:
    config = get_azure_openai_config(prefix="budg")
    return run_azure_openai_text(
        config,
        messages,
        max_tokens=max_tokens,
        temperature=temperature,
    )


def _call_ai_red_flags(summary: dict) -> list[dict]:
    system = """You are a senior credit risk analyst reviewing AR portfolio data for senior management.
Return ONLY a valid JSON array and nothing else.
Each object must contain exactly: severity, title, detail.
Allowed severity values: critical, warning, info.
Return 3 to 6 items ordered from most severe to least severe.
Use only the numbers in the input.
Focus on uninsured late AR, doubtful customers with large balances, weak current-quarter collections on large overdue accounts, provision increase drivers, focus list customers at risk, and low coverage."""

    user = f"""AR portfolio data:

{json.dumps(summary, indent=2)}

Return the JSON array only."""

    try:
        raw = _ai_chat_completion(
            [
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            max_tokens=1000,
            temperature=0.2,
        )
        if raw.startswith("```"):
            parts = raw.split("```")
            raw = parts[1] if len(parts) > 1 else raw
            if raw.startswith("json"):
                raw = raw[4:]
        flags = json.loads(raw.strip())
        return flags if isinstance(flags, list) else []
    except json.JSONDecodeError as e:
        return [{"severity": "info", "title": "Could not parse AI response", "detail": f"JSON error: {e}"}]
    except Exception as e:
        return [{"severity": "info", "title": "AI analysis unavailable", "detail": str(e)}]


def _render_red_flags(df: pd.DataFrame, selected_quarter: str):
    st.markdown("#### AI Risk Flags")

    with st.spinner("Scanning portfolio for risk signals..."):
        summary = _build_data_summary(df, selected_quarter)
        flags = _call_ai_red_flags(summary)

    if not flags:
        st.success("No significant risk flags detected.")
        return

    sev_map = {
        "critical": ("Critical", "error"),
        "warning": ("Warning", "warning"),
        "info": ("Info", "info"),
    }

    for flag in flags:
        severity = str(flag.get("severity", "info")).lower()
        label, box_type = sev_map.get(severity, ("Info", "info"))
        title = flag.get("title", "")
        detail = flag.get("detail", "")
        msg = f"**{label}: {title}**  \n{detail}"

        if box_type == "error":
            st.error(msg)
        elif box_type == "warning":
            st.warning(msg)
        else:
            st.info(msg)


def _status_donut(df: pd.DataFrame) -> go.Figure:
    if _STATUS not in df.columns:
        return go.Figure()

    grouped = df.groupby(_STATUS).apply(lambda group: _num(group, _AR_BAL).sum()).reset_index()
    grouped.columns = ["Status", "AR"]
    grouped = grouped[grouped["AR"] > 0].sort_values("AR", ascending=False)

    palette = ["#0f766e", "#0ea5e9", "#f59e0b", "#ef4444", "#14b8a6", "#64748b"]
    fig = go.Figure(
        go.Pie(
            labels=grouped["Status"],
            values=(grouped["AR"] / 1_000_000).round(2),
            hole=0.54,
            marker_colors=palette[: len(grouped)],
            textinfo="label+percent",
            textfont=dict(size=11),
            hovertemplate="%{label}: $%{value:.2f}M<extra></extra>",
        )
    )
    fig.update_layout(showlegend=False, **_CHART_LAYOUT)
    return fig


def _top_exposures(df: pd.DataFrame, n: int = 10):
    needed = [_CUST_NAME, _AR_BAL]
    for column in [_REGION, _STATUS, _FOCUS]:
        if column in df.columns:
            needed.append(column)

    view = df[needed].copy()
    view["AR ($M)"] = (_num(df, _AR_BAL) / 1_000_000).round(2)
    view = view.sort_values("AR ($M)", ascending=False).head(n).reset_index(drop=True)

    rename = {_CUST_NAME: "Customer", _REGION: "Region", _STATUS: "Status", _FOCUS: "Focus"}
    view = view.rename(columns={key: value for key, value in rename.items() if key in view.columns})
    cols = [column for column in ["Customer", "Region", "Status", "Focus", "AR ($M)"] if column in view.columns]
    view = view[cols]

    def _color_status(value):
        normalized = str(value).upper()
        if normalized == "GOOD":
            return "background-color:#e8f5e9;color:#245b2a"
        if normalized in ("REGULAR", "SUBSTANDARD"):
            return "background-color:#fff4d6;color:#8a4b08"
        if normalized in ("DOUBTFUL", "DOUB_INS_CLA"):
            return "background-color:#fdeaea;color:#9f1239"
        return ""

    if "Status" in view.columns:
        styler = view.style
        styled = styler.map(_color_status, subset=["Status"]) if hasattr(styler, "map") else styler.applymap(
            _color_status,
            subset=["Status"],
        )
    else:
        styled = view.style

    st.dataframe(
        styled,
        use_container_width=True,
        hide_index=True,
        height=min(38 * (n + 1) + 4, 420),
    )


def _call_ai_question(question: str, summary: dict) -> str:
    prompt = f"""You are a senior credit risk analyst.
Use only the portfolio data below and answer the management question clearly and briefly.
Include specific customer names, dollar amounts, and percentages where relevant.
Do not invent any numbers.

Portfolio data:
{json.dumps(summary, indent=2)}

Question: {question}"""

    try:
        return _ai_chat_completion([{"role": "user", "content": prompt}], max_tokens=500, temperature=0.2)
    except Exception as e:
        return f"Error: {e}"


def _render_chat(summary: dict):
    st.markdown("#### Ask your AR data")
    st.caption("Type any question about your portfolio and the copilot will answer from your actual numbers.")

    st.markdown("**Quick questions:**")
    question_cols = st.columns(3)
    example_questions = [
        "Which customers should we call first this week?",
        "What is our total uninsured exposure above 90 days?",
        "Which good-status customers have suspicious aging?",
    ]
    for idx, question in enumerate(example_questions):
        with question_cols[idx]:
            if st.button(question, key=f"quick_q_{idx}", use_container_width=True):
                st.session_state["bud_chat_input"] = question

    if "bud_chat_history" not in st.session_state:
        st.session_state["bud_chat_history"] = []

    question = st.chat_input("Ask anything about your AR portfolio...")
    if st.session_state.get("bud_chat_input"):
        question = st.session_state.pop("bud_chat_input")

    if question:
        with st.spinner("Analysing..."):
            answer = _call_ai_question(question, summary)
        st.session_state["bud_chat_history"].append({"question": question, "answer": answer})

    for item in reversed(st.session_state["bud_chat_history"]):
        with st.chat_message("user"):
            st.write(item["question"])
        with st.chat_message("assistant"):
            st.write(item["answer"])

    if st.session_state["bud_chat_history"] and st.button("Clear conversation", key="clear_chat"):
        st.session_state["bud_chat_history"] = []
        st.rerun()


def render_dashboard(bud_rows: pd.DataFrame, selected_quarter: str = "Q1"):
    if bud_rows is None or bud_rows.empty:
        st.info("No data to display. Upload a By_Customer file first.")
        return

    st.markdown(
        f"<div style='font-size:13px;color:var(--color-text-secondary);margin-bottom:0.5rem;'>"
        f"Starting quarter: <strong>{selected_quarter}</strong>"
        f" &nbsp;|&nbsp; {len(bud_rows):,} customer rows</div>",
        unsafe_allow_html=True,
    )
    st.divider()

    render_azure_openai_settings(prefix="budg")
    st.divider()

    _render_red_flags(bud_rows, selected_quarter)
    st.divider()

    col_donut, col_table = st.columns([1, 2], gap="large")

    with col_donut:
        st.markdown("**AR balance by customer status**")
        fig = _status_donut(bud_rows)
        if fig.data:
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        else:
            st.caption("No customer status data available.")

    with col_table:
        st.markdown("**Top exposures by AR balance**")
        _top_exposures(bud_rows)

    st.divider()

    summary = _build_data_summary(bud_rows, selected_quarter)
    _render_chat(summary)
