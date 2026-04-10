import html

import numpy as np
import pandas as pd
import streamlit as st

from common.quarter_utils import (
    QUARTER_ORDER,
    QUARTER_TAIL_LABELS_2026,
    build_customer_output_config,
)
from common.region_maps import classify_region

ZERO_QUARTER_CUSTOMER_KEYWORDS = ("MINDWARE", "AKLANIAT", "IFIX")
ZERO_COLLECTION_MAIN_ACCOUNTS = {"12302", "12304", "12306"}


def sanitize_colnames(df: pd.DataFrame) -> pd.DataFrame:
    """Decode HTML entities, remove NBSP, strip whitespace from column names."""

    def _clean(name: str) -> str:
        s = html.unescape(str(name))
        s = s.replace("\u00A0", " ").strip()
        return s

    return df.rename(columns={c: _clean(c) for c in df.columns})


def safe_to_datetime(series, column_name: str = "") -> pd.Series:
    """Parse dates with Streamlit error reporting for bad rows."""
    if series is None:
        return pd.to_datetime(pd.Series([], dtype="object"), errors="coerce")
    series_str = series.astype(str).str.strip().replace({"nan": np.nan, "": np.nan})
    dt = pd.to_datetime(series_str, errors="coerce")
    bad = dt.isna() & series_str.notna()
    if bad.any():
        st.error(f"âš ï¸ Invalid datetime values in column '{column_name}':")
        for idx, val in series_str[bad].items():
            st.write(f"Row {idx + 16} â†’ {val}")
    return dt


def process_ar_file(file):
    excel = pd.ExcelFile(file, engine="openpyxl")
    sheet = excel.sheet_names[0]

    header_block = pd.read_excel(excel, sheet_name=sheet, header=None, nrows=14)
    as_on_date = header_block.iloc[13, 1]
    if pd.isna(as_on_date):
        raise ValueError("Cell B14 must contain 'As on Date'")
    as_on_date = pd.to_datetime(str(as_on_date).strip(), errors="coerce")
    if pd.isna(as_on_date):
        raise ValueError("As on Date in B14 is not a valid date")

    df = pd.read_excel(excel, sheet_name=sheet, header=15, dtype=str)
    df = sanitize_colnames(df)
    original_order = list(df.columns)

    ar_balance = pd.to_numeric(df.get("Ar Balance"), errors="coerce").fillna(0)
    doc_date = safe_to_datetime(df.get("Document Date"), "Document Date")
    due_date = safe_to_datetime(df.get("Document Due Date"), "Document Due Date")
    if len(doc_date) != len(df):
        doc_date = pd.to_datetime(pd.Series([pd.NaT] * len(df)), errors="coerce")
    if len(due_date) != len(df):
        due_date = pd.to_datetime(pd.Series([pd.NaT] * len(df)), errors="coerce")

    doc_date_filled = doc_date.fillna(as_on_date)
    due_date_filled = due_date.fillna(as_on_date)

    ageing_days = (as_on_date - doc_date_filled).dt.days
    overdue_days = (as_on_date - due_date_filled).dt.days
    overdue_num = pd.to_numeric(overdue_days, errors="coerce").fillna(0)

    if "Cust Region" in df.columns and "Cust Code" in df.columns:
        region_series = classify_region(df["Cust Region"], df["Cust Code"])
    elif "Cust Region" in df.columns:
        region_series = classify_region(df["Cust Region"])
    elif "Cust Code" in df.columns:
        empty_region = pd.Series([""] * len(df), index=df.index)
        region_series = classify_region(empty_region, df["Cust Code"])
    else:
        region_series = pd.Series(["KSA"] * len(df), index=df.index)

    if "Customer Status" in df.columns:
        updated_status = df["Customer Status"].fillna("").replace("", "SUBSTANDARD")
    else:
        updated_status = pd.Series(["SUBSTANDARD"] * len(df), index=df.index)

    conditions = [
        ar_balance < 0,
        overdue_num < 0,
        (overdue_num >= 0) & (overdue_num <= 30),
        (overdue_num > 30) & (overdue_num <= 60),
        (overdue_num > 60) & (overdue_num <= 90),
        (overdue_num > 90) & (overdue_num <= 120),
        (overdue_num > 120) & (overdue_num <= 150),
        overdue_num > 150,
    ]
    choices = [
        "On account",
        "Not Due",
        "Aging 1 to 30",
        "Aging 31 to 60",
        "Aging 61 to 90",
        "Aging 91 to 120",
        "Aging 121 to 150",
        "Aging >=151",
    ]
    aging_bracket_label = np.select(conditions, choices, default="")

    invoice_value = ar_balance.clip(lower=0)
    on_account_amount = ar_balance.clip(upper=0)
    not_due_amount = np.where(overdue_num > 0, 0, invoice_value)

    BP, BK = invoice_value, overdue_num
    amt_ge151 = np.where(BK > 150, BP, 0)
    amt_121_150 = np.where(BK > 120, BP, 0) - amt_ge151
    amt_91_120 = np.where(BK > 90, BP, 0) - amt_121_150 - amt_ge151
    amt_61_90 = np.where(BK > 60, BP, 0) - amt_91_120 - amt_121_150 - amt_ge151
    amt_31_60 = np.where(BK > 30, BP, 0) - amt_61_90 - amt_91_120 - amt_121_150 - amt_ge151
    amt_1_30 = np.where(BK >= 0, BP, 0) - amt_31_60 - amt_61_90 - amt_91_120 - amt_121_150 - amt_ge151
    for arr in [amt_1_30, amt_31_60, amt_61_90, amt_91_120, amt_121_150, amt_ge151]:
        np.maximum(arr, 0, out=arr)

    df["Ageing (Days)"] = ageing_days
    df["Overdue days (Days)"] = overdue_num
    df["Region (Derived)"] = region_series
    df["Aging Bracket (Label)"] = aging_bracket_label
    df["Updated Status"] = updated_status
    df["Invoice Value (Derived)"] = invoice_value
    df["On Account (Derived)"] = on_account_amount
    df["Not Due (Derived)"] = not_due_amount
    df["Aging 1 to 30 (Amount)"] = amt_1_30
    df["Aging 31 to 60 (Amount)"] = amt_31_60
    df["Aging 61 to 90 (Amount)"] = amt_61_90
    df["Aging 91 to 120 (Amount)"] = amt_91_120
    df["Aging 121 to 150 (Amount)"] = amt_121_150
    df["Aging >=151 (Amount)"] = amt_ge151
    df["Ageing > 365 (Amt)"] = np.where(df["Ageing (Days)"] > 365, ar_balance, 0)
    df["Ar Balance (Copy)"] = ar_balance

    appended_block = [
        "Ageing (Days)",
        "Overdue days (Days)",
        "Region (Derived)",
        "Ar Balance (Copy)",
        "Aging Bracket (Label)",
        "Updated Status",
        "Invoice Value (Derived)",
        "On Account (Derived)",
        "Not Due (Derived)",
        "Aging 1 to 30 (Amount)",
        "Aging 31 to 60 (Amount)",
        "Aging 61 to 90 (Amount)",
        "Aging 91 to 120 (Amount)",
        "Aging 121 to 150 (Amount)",
        "Aging >=151 (Amount)",
        "Ageing > 365 (Amt)",
    ]
    for col in appended_block:
        if col not in df.columns:
            df[col] = ""
    df = df[list(original_order) + appended_block]
    df.attrs["as_on_date"] = as_on_date
    return df


def customer_summary(df, selected_quarter="Q1"):
    df = sanitize_colnames(df)
    out = df.copy()
    out = out.loc[:, ~out.columns.duplicated(keep="last")]
    cfg = build_customer_output_config(selected_quarter)

    out["Cust Code"] = out.get("Cust Code", "").astype(str).str.strip()
    out["Main Ac"] = out.get("Main Ac", "").fillna("").astype(str).str.strip()
    out["Cust Name"] = out.get("Cust Name", "").fillna("").astype(str).str.strip()

    if "Region" not in out.columns:
        if "Region (Derived)" in out.columns:
            out["Region"] = out["Region (Derived)"]
        elif "Cust Region" in out.columns:
            out["Region"] = out["Cust Region"]
        else:
            out["Region"] = ""

    if "Not Due Amount" not in out.columns:
        raise ValueError("Template error: missing column 'Not Due Amount'. Upload the standard AR Backlog template.")

    for c in [
        "On Account (Derived)",
        "Aging 1 to 30 (Amount)",
        "Aging 31 to 60 (Amount)",
        "Aging 61 to 90 (Amount)",
        "Aging 91 to 120 (Amount)",
        "Aging 121 to 150 (Amount)",
        "Aging >=151 (Amount)",
        "Ageing > 365 (Amt)",
        "Overdue days (Days)",
        "Not Due Amount",
        "Ar Balance (Copy)",
    ]:
        out[c] = pd.to_numeric(out.get(c, 0), errors="coerce").fillna(0)

    if "Document Due Date" in out.columns:
        due_raw = out["Document Due Date"].astype(str).str.strip()
        due_raw = due_raw.replace("\u00A0", " ", regex=False)
        due_raw = due_raw.str.replace(r"[^\x00-\x7F]", "", regex=True)
        due_raw = due_raw.str.replace(r"\s+\d{2}:\d{2}:\d{2}$", "", regex=True)
        due_dt = pd.to_datetime(due_raw, errors="coerce")
    else:
        due_dt = pd.Series([pd.NaT] * len(out))


    yr = 2026
    # Define the period mapping for each quarter selection
    period_map = {
        "Q1": [
            ("Q1-2026", (pd.Timestamp(yr, 1, 1), pd.Timestamp(yr, 3, 15))),
            ("Q1 tail", (pd.Timestamp(yr, 3, 16), pd.Timestamp(yr, 3, 31))),
            ("Q2-2026", (pd.Timestamp(yr, 3, 16), pd.Timestamp(yr, 6, 15))),
            ("Q2 tail", (pd.Timestamp(yr, 6, 16), pd.Timestamp(yr, 6, 30))),
            ("Q3-2026", (pd.Timestamp(yr, 6, 16), pd.Timestamp(yr, 9, 15))),
            ("Q3 tail", (pd.Timestamp(yr, 9, 16), pd.Timestamp(yr, 9, 30))),
            ("Q4-2026", (pd.Timestamp(yr, 9, 16), pd.Timestamp(yr, 12, 15))),
            ("Q4 tail", (pd.Timestamp(yr, 12, 16), pd.Timestamp(yr, 12, 31))),
            ("2027", (pd.Timestamp(2027, 12, 16), pd.Timestamp(2027, 12, 31))),
            ("2028", (pd.Timestamp(2028, 1, 1), pd.Timestamp(2028, 12, 31))),
            ("2029", (pd.Timestamp(2029, 1, 1), pd.Timestamp(2029, 12, 31))),
            ("2030", (pd.Timestamp(2030, 1, 1), pd.Timestamp(2030, 12, 31))),
        ],
        "Q2": [
            ("Q2-2026", (pd.Timestamp(yr, 4, 1), pd.Timestamp(yr, 6, 15))),
            ("Q2 tail", (pd.Timestamp(yr, 6, 16), pd.Timestamp(yr, 6, 30))),
            ("Q3-2026", (pd.Timestamp(yr, 6, 16), pd.Timestamp(yr, 9, 15))),
            ("Q3 tail", (pd.Timestamp(yr, 9, 16), pd.Timestamp(yr, 9, 30))),
            ("Q4-2026", (pd.Timestamp(yr, 9, 16), pd.Timestamp(yr, 12, 15))),
            ("Q4 tail", (pd.Timestamp(yr, 12, 16), pd.Timestamp(yr, 12, 31))),
            ("2027", (pd.Timestamp(2027, 12, 16), pd.Timestamp(2027, 12, 31))),
            ("2028", (pd.Timestamp(2028, 1, 1), pd.Timestamp(2028, 12, 31))),
            ("2029", (pd.Timestamp(2029, 1, 1), pd.Timestamp(2029, 12, 31))),
            ("2030", (pd.Timestamp(2030, 1, 1), pd.Timestamp(2030, 12, 31))),
        ],
        "Q3": [
            ("Q3-2026", (pd.Timestamp(yr, 7, 1), pd.Timestamp(yr, 9, 15))),
            ("Q3 tail", (pd.Timestamp(yr, 9, 16), pd.Timestamp(yr, 9, 30))),
            ("Q4-2026", (pd.Timestamp(yr, 9, 16), pd.Timestamp(yr, 12, 15))),
            ("Q4 tail", (pd.Timestamp(yr, 12, 16), pd.Timestamp(yr, 12, 31))),
            ("2027", (pd.Timestamp(2027, 12, 16), pd.Timestamp(2027, 12, 31))),
            ("2028", (pd.Timestamp(2028, 1, 1), pd.Timestamp(2028, 12, 31))),
            ("2029", (pd.Timestamp(2029, 1, 1), pd.Timestamp(2029, 12, 31))),
            ("2030", (pd.Timestamp(2030, 1, 1), pd.Timestamp(2030, 12, 31))),
        ],
        "Q4": [
            ("Q4-2026", (pd.Timestamp(yr, 10, 1), pd.Timestamp(yr, 12, 15))),
            ("Q4 tail", (pd.Timestamp(yr, 12, 16), pd.Timestamp(yr, 12, 31))),
            ("2027", (pd.Timestamp(2027, 12, 16), pd.Timestamp(2027, 12, 31))),
            ("2028", (pd.Timestamp(2028, 1, 1), pd.Timestamp(2028, 12, 31))),
            ("2029", (pd.Timestamp(2029, 1, 1), pd.Timestamp(2029, 12, 31))),
            ("2030", (pd.Timestamp(2030, 1, 1), pd.Timestamp(2030, 12, 31))),
        ],
    }

    # Quarter/tail/year allocations should only use collectible positive balances.
    blocked_customer = out["Cust Name"].str.upper().str.contains(
        "|".join(ZERO_QUARTER_CUSTOMER_KEYWORDS), na=False
    )
    blocked_main_account = out["Main Ac"].isin(ZERO_COLLECTION_MAIN_ACCOUNTS)
    # Block if Updated Status is not GOOD, Regular, or Substandard (case-insensitive)
    allowed_statuses = ["GOOD", "REGULAR", "SUBSTANDARD"]
    blocked_status = ~out["Updated Status"].str.upper().isin(allowed_statuses)
    quarter_amount = out["Ar Balance (Copy)"].clip(lower=0).where(~(blocked_customer | blocked_main_account | blocked_status), 0)

    # Assign values to each period column strictly as per the mapping
    for col, (start, end) in period_map[selected_quarter]:
        out[col] = np.where((due_dt >= start) & (due_dt <= end), quarter_amount, 0)

    # Aggregation and output columns
    agg_map = {
        "Cust Name": ("Cust Name", "first"),
        "Cust Region": ("Cust Region", "first"),
        "Region": ("Region", "first"),
        "Updated Status": ("Updated Status", "first"),
        "On Account (Derived)": ("On Account (Derived)", "sum"),
        "Not Due Amount": ("Not Due Amount", "sum"),
        "AR Balance": ("Ar Balance (Copy)", "sum"),
        "Overdue days (Days)": ("Overdue days (Days)", "sum"),
        "Aging 1 to 30 (Amount)": ("Aging 1 to 30 (Amount)", "sum"),
        "Aging 31 to 60 (Amount)": ("Aging 31 to 60 (Amount)", "sum"),
        "Aging 61 to 90 (Amount)": ("Aging 61 to 90 (Amount)", "sum"),
        "Aging 91 to 120 (Amount)": ("Aging 91 to 120 (Amount)", "sum"),
        "Aging 121 to 150 (Amount)": ("Aging 121 to 150 (Amount)", "sum"),
        "Aging >=151 (Amount)": ("Aging >=151 (Amount)", "sum"),
        "Ageing > 365 (Amt)": ("Ageing > 365 (Amt)", "sum"),
    }
    for col, _ in period_map[selected_quarter]:
        agg_map[col] = (col, "sum")

    grouped = out.groupby(["Cust Code", "Main Ac"], as_index=False).agg(**agg_map)

    amount_buckets = [
        "Aging 1 to 30 (Amount)",
        "Aging 31 to 60 (Amount)",
        "Aging 61 to 90 (Amount)",
        "Aging 91 to 120 (Amount)",
        "Aging 121 to 150 (Amount)",
        "Aging >=151 (Amount)",
    ]
    present = [c for c in amount_buckets if c in grouped.columns]
    grouped["Overdue days (Days)"] = grouped[present].sum(axis=1) if present else 0

    rename_final = {
        "On Account (Derived)": "On account",
        "Not Due Amount": "Not Due",
        "AR Balance": "Ar Balance",
        "Overdue days (Days)": "Overdue",
        "Aging 1 to 30 (Amount)": "Aging 1 to 30",
        "Aging 31 to 60 (Amount)": "Aging 31 to 60",
        "Aging 61 to 90 (Amount)": "Aging 61 to 90",
        "Aging 91 to 120 (Amount)": "Aging 91 to 120",
        "Aging 121 to 150 (Amount)": "Aging 121 to 150",
        "Aging >=151 (Amount)": "Aging >=151",
        "Ageing > 365 (Amt)": "Ageing > 365",
    }
    grouped = grouped.rename(columns=rename_final)

    dynamic_manual_cols = [
        cfg["percent_label"],
        cfg["actual_label"],
        cfg["remaining_label"],
        cfg["to_add_label"],
        cfg["forecast_label"],
    ]
    for col in dynamic_manual_cols:
        if col not in grouped.columns:
            grouped[col] = 0

    # Strict column order for output
    final_order = [
        "Cust Code",
        "Cust Name",
        "Main Ac",
        "Cust Region",
        "Region",
        "Updated Status",
        "On account",
        "Not Due",
        "Ar Balance",
        "Overdue",
        "Aging 1 to 30",
        "Aging 31 to 60",
        "Aging 61 to 90",
        "Aging 91 to 120",
        "Aging 121 to 150",
        "Aging >=151",
        "Ageing > 365",
    ]
    for col, _ in period_map[selected_quarter]:
        final_order.append(col)
    for c in final_order:
        if c not in grouped.columns:
            grouped[c] = 0
    grouped.attrs["selected_quarter"] = selected_quarter
    return grouped[final_order]


def invoice_summary(df):
    work = sanitize_colnames(df).copy()
    work = work.loc[:, ~work.columns.duplicated(keep="last")]

    if "Region" not in work.columns:
        if "Region (Derived)" in work.columns:
            work["Region"] = work["Region (Derived)"]
        elif "Cust Region" in work.columns:
            work["Region"] = work["Cust Region"]
        else:
            work["Region"] = ""

    if "Ar Balance (Copy)" in work.columns:
        ar_source = "Ar Balance (Copy)"
    elif "Ar Balance" in work.columns:
        ar_source = "Ar Balance"
    else:
        work["Ar Balance (Copy)"] = 0
        ar_source = "Ar Balance (Copy)"

    numeric_map = {
        "Ageing (Days)": "Ageing",
        "Overdue days (Days)": "Overdue days",
        "On Account (Derived)": "On Account",
        "Not Due (Derived)": "Not Due",
        ar_source: "Ar Balance",
        "Aging 1 to 30 (Amount)": "Aging 1 to 30",
        "Aging 31 to 60 (Amount)": "Aging 31 to 60",
        "Aging 61 to 90 (Amount)": "Aging 61 to 90",
        "Aging 91 to 120 (Amount)": "Aging 91 to 120",
        "Aging 121 to 150 (Amount)": "Aging 121 to 150",
        "Aging >=151 (Amount)": "Aging >=151",
    }

    out = pd.DataFrame()

    def copy_if_exists(src, dst):
        out[dst] = work[src] if src in work.columns else ""

    for src in [
        "Cust Code",
        "Cust Name",
        "Main Ac",
        "Cust Region",
        "Region",
        "Document Number",
        "Document Date",
        "Document Due Date",
        "Payment Terms",
        "Brand",
        "Total Insurance Limit",
        "LC & BG Guarantee",
        "SO No",
        "LPO No",
    ]:
        copy_if_exists(src, src)

    for src, dst in numeric_map.items():
        out[dst] = pd.to_numeric(work.get(src, 0), errors="coerce").fillna(0)

    final_order = [
        "Cust Code",
        "Cust Name",
        "Main Ac",
        "Cust Region",
        "Region",
        "Document Number",
        "Document Date",
        "Document Due Date",
        "Ageing",
        "Overdue days",
        "Payment Terms",
        "On Account",
        "Not Due",
        "Ar Balance",
        "Aging 1 to 30",
        "Aging 31 to 60",
        "Aging 61 to 90",
        "Aging 91 to 120",
        "Aging 121 to 150",
        "Aging >=151",
        "Brand",
        "Total Insurance Limit",
        "LC & BG Guarantee",
        "SO No",
        "LPO No",
    ]
    for c in final_order:
        if c not in out.columns:
            out[c] = ""
    return out[final_order]
