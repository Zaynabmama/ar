"""
budg/bud2026_dashboard.py

Management dashboard for BUD2026.
- AI-powered red flag detector (uses Claude Haiku via Anthropic API)
- AR balance by customer status (donut)
- Top exposures by AR balance (table)

Requires: plotly, anthropic
  pip install plotly anthropic
"""

import json
import os

import anthropic
import groq
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ── Column constants (match HEADERS_BUD2026 exactly) ──────────────────────
_AR_BAL          = " AR\nBalance"
_INSURANCE       = "Insurance"
_A121_150        = "Aging\n121 to 150"
_AGE_GE_151      = "Aging\n>=151"
_PROV_INCL_INS   = "Provision after collection including Insurance/BG/LC"
_OPEN_PROV       = "AR Provision at\n31/08/2025"
_REGION          = "Sales Budget region"
_STATUS          = "Customer Status"
_CUST_NAME       = "Cust Name"
_FOCUS           = "Focus List"
_FC_Q1           = "Collections FC\n31/03/2026"

_CHART_LAYOUT = dict(
    margin=dict(l=0, r=0, t=8, b=0),
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(size=12),
)


# ── Helpers ────────────────────────────────────────────────────────────────

def _num(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series(0.0, index=df.index)
    return pd.to_numeric(df[col], errors="coerce").fillna(0.0)


def _fmt_m(val: float, decimals: int = 1) -> str:
    if abs(val) >= 1_000_000:
        return f"${val / 1_000_000:.{decimals}f}M"
    if abs(val) >= 1_000:
        return f"${val / 1_000:.{decimals}f}K"
    return f"${val:,.0f}"


# ── Build compact data summary to send to AI ──────────────────────────────

def _build_data_summary(df: pd.DataFrame, selected_quarter: str) -> dict:
    """
    Condenses bud_rows into a small JSON dict that fits cheaply in a Haiku prompt.
    All money values are in $M rounded to 2 decimals.
    """
    total_ar   = float(_num(df, _AR_BAL).sum())
    total_prov = float(_num(df, _PROV_INCL_INS).sum())
    open_prov  = float(_num(df, _OPEN_PROV).sum())
    prov_diff  = total_prov - open_prov
    q1_fc      = float(_num(df, _FC_Q1).sum())
    coverage   = round(total_prov / total_ar * 100, 1) if total_ar else 0
    late_ar    = float((_num(df, _A121_150) + _num(df, _AGE_GE_151)).sum())
    late_pct   = round(late_ar / total_ar * 100, 1) if total_ar else 0

    # Status split
    status_split = {}
    if _STATUS in df.columns:
        for s, g in df.groupby(_STATUS):
            status_split[str(s)] = round(float(_num(g, _AR_BAL).sum()) / 1_000_000, 2)

    # Region split
    region_split = {}
    if _REGION in df.columns:
        for r, g in df.groupby(_REGION):
            region_split[str(r)] = round(float(_num(g, _AR_BAL).sum()) / 1_000_000, 2)

    # Enrich working copy once
    w = df.copy()
    w["__ar"]   = _num(w, _AR_BAL)
    w["__ins"]  = _num(w, _INSURANCE)
    w["__late"] = _num(w, _A121_150) + _num(w, _AGE_GE_151)
    w["__q1fc"] = _num(w, _FC_Q1)
    w["__prov"] = _num(w, _PROV_INCL_INS)

    # Top 10 by AR
    top10 = []
    for _, row in w.nlargest(10, "__ar").iterrows():
        top10.append({
            "name":        str(row.get(_CUST_NAME, "")),
            "region":      str(row.get(_REGION, "")),
            "status":      str(row.get(_STATUS, "")),
            "focus_list":  str(row.get(_FOCUS, "")),
            "ar_m":        round(float(row["__ar"]) / 1_000_000, 2),
            "insurance_m": round(float(row["__ins"]) / 1_000_000, 2),
            "late_ar_m":   round(float(row["__late"]) / 1_000_000, 2),
            "q1_fc_m":     round(float(row["__q1fc"]) / 1_000_000, 2),
            "provision_m": round(float(row["__prov"]) / 1_000_000, 2),
        })

    # Uninsured late AR (top 5 worst)
    uninsured = w[(w["__late"] > 0) & (w["__ins"] == 0)].nlargest(5, "__late")
    uninsured_list = [
        {
            "name":      str(r.get(_CUST_NAME, "")),
            "late_ar_m": round(float(r["__late"]) / 1_000_000, 2),
            "status":    str(r.get(_STATUS, "")),
            "region":    str(r.get(_REGION, "")),
        }
        for _, r in uninsured.iterrows()
    ]

    # Focus list customers
    focus_list = []
    if _FOCUS in df.columns:
        focus_df = w[w[_FOCUS].astype(str).str.strip().str.upper().isin(["YES", "Y", "X", "1", "TRUE"])]
        for _, row in focus_df.iterrows():
            focus_list.append({
                "name":    str(row.get(_CUST_NAME, "")),
                "ar_m":    round(float(row["__ar"]) / 1_000_000, 2),
                "q1_fc_m": round(float(row["__q1fc"]) / 1_000_000, 2),
                "status":  str(row.get(_STATUS, "")),
            })

    return {
        "selected_quarter":    selected_quarter,
        "total_ar_m":          round(total_ar / 1_000_000, 2),
        "total_provision_m":   round(total_prov / 1_000_000, 2),
        "opening_provision_m": round(open_prov / 1_000_000, 2),
        "provision_change_m":  round(prov_diff / 1_000_000, 2),
        "q1_fc_collections_m": round(q1_fc / 1_000_000, 2),
        "coverage_pct":        coverage,
        "late_ar_pct":         late_pct,
        "late_ar_m":           round(late_ar / 1_000_000, 2),
        "customer_count":      int(len(df)),
        "status_breakdown_m":  status_split,
        "region_breakdown_m":  region_split,
        "top_10_customers":    top10,
        "uninsured_late_ar":   uninsured_list,
        "focus_list_customers":focus_list,
    }


# ── Call Claude Haiku ──────────────────────────────────────────────────────
def _call_ai_red_flags(summary: dict) -> list[dict]:
    try:
        api_key = st.secrets["GROQ_API_KEY"]
    except Exception:
        api_key = "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"
    if not api_key:
        return [{"severity": "info", "title": "API key not configured",
                 "detail": "Add GROQ_API_KEY to .streamlit/secrets.toml."}]


    client = groq.Groq(api_key=api_key)
    system = """You are a senior credit risk analyst reviewing AR (accounts receivable) data.
You receive a JSON portfolio summary and identify the most important red flags for senior management.

STRICT OUTPUT RULES:
- Return ONLY a valid JSON array. Nothing else — no prose, no markdown fences.
- Each element must have exactly: "severity" (critical/warning/info), "title" (≤8 words), "detail" (1-2 sentences with specific numbers from the data).
- Return 3 to 6 flags, ordered most to least severe.
- Be specific: name customers, state dollar amounts and percentages from the data.
- Never invent numbers not in the input.
- Focus areas: uninsured late AR, doubtful customers with large balances, zero Q1 FC on large overdue accounts, provision increase drivers, focus list customers at risk, low coverage ratio."""

    user = f"""AR portfolio data for {summary['selected_quarter']} 2026:

{json.dumps(summary, indent=2)}

Return the JSON array of red flags only."""

    try:
        resp = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            max_tokens=1000,
            temperature=0.2,
        )
        raw = resp.choices[0].message.content.strip()

        # Strip accidental markdown fences
        if raw.startswith("```"):
            parts = raw.split("```")
            raw = parts[1] if len(parts) > 1 else raw
            if raw.startswith("json"):
                raw = raw[4:]
        flags = json.loads(raw.strip())
        return flags if isinstance(flags, list) else []
    except json.JSONDecodeError as e:
        return [{"severity": "info", "title": "Could not parse AI response",
                 "detail": f"JSON error: {e}"}]
    except Exception as e:
        return [{"severity": "info", "title": "AI analysis unavailable",
                 "detail": str(e)}]


# ── Render red flags section ───────────────────────────────────────────────

def _render_red_flags(df: pd.DataFrame, selected_quarter: str):
    st.markdown("#### AI Risk Flags")

    with st.spinner("Scanning portfolio for risk signals..."):
        summary = _build_data_summary(df, selected_quarter)
        flags   = _call_ai_red_flags(summary)

    if not flags:
        st.success("No significant risk flags detected.")
        return

    sev_map = {
        "critical": ("🔴", "error"),
        "warning":  ("🟡", "warning"),
        "info":     ("🔵", "info"),
    }

    for flag in flags:
        sev             = flag.get("severity", "info")
        icon, box_type  = sev_map.get(sev, ("🔵", "info"))
        title           = flag.get("title", "")
        detail          = flag.get("detail", "")
        msg             = f"**{icon} {title}**  \n{detail}"

        if box_type == "error":
            st.error(msg)
        elif box_type == "warning":
            st.warning(msg)
        else:
            st.info(msg)


# ── Status donut ───────────────────────────────────────────────────────────

def _status_donut(df: pd.DataFrame) -> go.Figure:
    if _STATUS not in df.columns:
        return go.Figure()

    grp = (
        df.groupby(_STATUS)
        .apply(lambda g: _num(g, _AR_BAL).sum())
        .reset_index()
    )
    grp.columns = ["Status", "AR"]
    grp = grp[grp["AR"] > 0].sort_values("AR", ascending=False)

    palette = ["#639922", "#378ADD", "#EF9F27", "#E24B4A", "#7F77DD", "#5DCAA5"]
    fig = go.Figure(go.Pie(
        labels=grp["Status"],
        values=(grp["AR"] / 1_000_000).round(2),
        hole=0.54,
        marker_colors=palette[: len(grp)],
        textinfo="label+percent",
        textfont=dict(size=11),
        hovertemplate="%{label}: $%{value:.2f}M<extra></extra>",
    ))
    fig.update_layout(showlegend=False, **_CHART_LAYOUT)
    return fig


# ── Top exposures table ────────────────────────────────────────────────────

def _top_exposures(df: pd.DataFrame, n: int = 10):
    needed = [_CUST_NAME, _AR_BAL]
    for c in [_REGION, _STATUS, _FOCUS]:
        if c in df.columns:
            needed.append(c)

    view = df[needed].copy()
    view["AR ($M)"] = (_num(df, _AR_BAL) / 1_000_000).round(2)
    view = view.sort_values("AR ($M)", ascending=False).head(n).reset_index(drop=True)

    rename = {_CUST_NAME: "Customer", _REGION: "Region",
              _STATUS: "Status", _FOCUS: "Focus"}
    view = view.rename(columns={k: v for k, v in rename.items() if k in view.columns})
    cols = [c for c in ["Customer", "Region", "Status", "Focus", "AR ($M)"] if c in view.columns]
    view = view[cols]

    def _color_status(val):
        v = str(val).upper()
        if v == "GOOD":
            return "background-color:#EAF3DE;color:#3B6D11"
        if v in ("REGULAR", "SUBSTANDARD"):
            return "background-color:#FAEEDA;color:#854F0B"
        if v in ("DOUBTFUL", "DOUB_INS_CLA"):
            return "background-color:#FCEBEB;color:#A32D2D"
        return ""

    styled = (
        view.style.applymap(_color_status, subset=["Status"])
        if "Status" in view.columns else view.style
    )
    st.dataframe(
        styled,
        use_container_width=True,
        hide_index=True,
        height=min(38 * (n + 1) + 4, 420),
    )


# ── AI chat ───────────────────────────────────────────────────────────────

def _call_ai_question(question: str, summary: dict) -> str:
    """
    Send a free-text question + data summary to Groq.
    Returns a plain English answer string.
    """
    try:
        api_key = st.secrets["GROQ_API_KEY"]
    except Exception:
        api_key = "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"


    try:
        client = groq.Groq(api_key=api_key)

        prompt = f"""You are a senior credit risk analyst. You have access to the following AR portfolio data:

{json.dumps(summary, indent=2)}

Answer the following question from senior management. Be specific — use customer names, dollar amounts, and percentages from the data above. Keep the answer concise (3-5 sentences max). Do not make up numbers that are not in the data.

Question: {question}"""

        response = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=400,
            temperature=0.2,
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Error: {e}"


def _render_chat(summary: dict):
    st.markdown("#### Ask your AR data")
    st.caption("Type any question about your portfolio — the AI reads your actual numbers")

    # Example questions as quick buttons
    st.markdown("**Quick questions:**")
    q_cols = st.columns(3)
    example_questions = [
        "Which customers should we call first this week?",
        "What is our total uninsured exposure above 90 days?",
        "Which good-status customers have suspicious aging?",
    ]
    for i, q in enumerate(example_questions):
        with q_cols[i]:
            if st.button(q, key=f"quick_q_{i}", use_container_width=True):
                st.session_state["bud_chat_input"] = q

    # Chat history stored in session state
    if "bud_chat_history" not in st.session_state:
        st.session_state["bud_chat_history"] = []

    # Input box
    question = st.chat_input("Ask anything about your AR portfolio...")

    # Also handle quick question buttons
    if "bud_chat_input" in st.session_state and st.session_state["bud_chat_input"]:
        question = st.session_state.pop("bud_chat_input")

    if question:
        with st.spinner("Analysing..."):
            answer = _call_ai_question(question, summary)
        st.session_state["bud_chat_history"].append({
            "question": question,
            "answer": answer,
        })

    # Display history — most recent first
    for item in reversed(st.session_state["bud_chat_history"]):
        with st.chat_message("user"):
            st.write(item["question"])
        with st.chat_message("assistant"):
            st.write(item["answer"])

    # Clear button
    if st.session_state["bud_chat_history"]:
        if st.button("Clear conversation", key="clear_chat"):
            st.session_state["bud_chat_history"] = []
            st.rerun()


# ── Main entry point ───────────────────────────────────────────────────────

def render_dashboard(bud_rows: pd.DataFrame, selected_quarter: str = "Q1"):
    """
    Call from ui_new_bud2026.py after bud_rows is computed:

        from budg.bud2026_dashboard import render_dashboard
        render_dashboard(bud_rows, selected_quarter=selected_quarter)
    """
    if bud_rows is None or bud_rows.empty:
        st.info("No data to display. Upload a By_Customer file first.")
        return

    st.markdown(
        f"<div style='font-size:13px;color:var(--color-text-secondary);margin-bottom:0.5rem;'>"
        f"Starting quarter: <strong>{selected_quarter}</strong>"
        f" &nbsp;·&nbsp; {len(bud_rows):,} customer rows</div>",
        unsafe_allow_html=True,
    )
    st.divider()

    # ── AI red flags ──────────────────────────────────────────────────────
    _render_red_flags(bud_rows, selected_quarter)

    st.divider()

    # ── Charts: donut + top exposures ─────────────────────────────────────
    col_donut, col_table = st.columns([1, 2], gap="large")

    with col_donut:
        st.markdown("**AR balance by customer status**")
        fig = _status_donut(bud_rows)
        if fig.data:
            st.plotly_chart(fig, use_container_width=True,
                            config={"displayModeBar": False})
        else:
            st.caption("No customer status data available.")

    with col_table:
        st.markdown("**Top exposures by AR balance**")
        _top_exposures(bud_rows)

    st.divider()

    # ── AI chat ───────────────────────────────────────────────────────────
    # Build summary once and reuse for chat (same data as red flags)
    summary = _build_data_summary(bud_rows, selected_quarter)
    _render_chat(summary)