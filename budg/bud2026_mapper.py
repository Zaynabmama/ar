# budg/bud2026_mapper.py
import pandas as pd
import numpy as np

from budg.bud2026_headers import QUARTER_COLLECTION_HEADERS
from common.identifier_utils import normalize_excel_identifier_series
from common.quarter_utils import QUARTER_ORDER, build_customer_output_config

try:
    from common.region_maps import classify_region
except Exception:
    classify_region = None

# ====================== HELPERS ======================

def _series_or_empty(df: pd.DataFrame, col: str) -> pd.Series:
    """Return column as Series if exists, else empty string Series"""
    if col in df.columns:
        return df[col]
    return pd.Series([""] * len(df), index=df.index)

def _num(df: pd.DataFrame, col: str) -> pd.Series:
    """Coerce to numeric; missing or invalid -> 0"""
    if not col or col not in df.columns:
        return pd.Series([0.0] * len(df), index=df.index, dtype="float64")
    return pd.to_numeric(df[col], errors="coerce").fillna(0.0)

def _derive_sales_budget_region(df_cust: pd.DataFrame) -> pd.Series:
    """Derive 'Sales Budget region' robustly"""
    if "Region" in df_cust.columns:
        reg = df_cust["Region"].fillna("").astype(str)
        if reg.str.strip().any():
            return reg
    if classify_region is not None and "Cust Region" in df_cust.columns:
        cust_code = df_cust.get("Cust Code", None)
        derived = classify_region(df_cust["Cust Region"], cust_code)
        return derived.fillna("")
    return pd.Series([""] * len(df_cust), index=df_cust.index)

def _first_present(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """First matching column, ignoring whitespace differences (By_Customer
    writes some headers with embedded newlines)."""
    normalized = {"".join(str(c).split()).lower(): c for c in df.columns}
    for c in candidates:
        hit = normalized.get("".join(c.split()).lower())
        if hit is not None:
            return hit
    return None

# By_Customer columns feeding the K-O Not Due breakdown (same source columns
# as the provision tool; matching ignores whitespace)
NOT_DUE_BREAKDOWN_COLS = {
    "Not Due\n0-30 days": "Not Due 0-30 days",
    "Not Due\n31-60 days": "Not Due 31-60 days",
    "Not Due\n61-90 days": "Not Due 61-90 days",
    "Not Due\n91-180 days": "Not Due 91-180 days",
    "Not Due\n180+ days": "Not Due 180+ days",
}

# Tool 1 (Orion) leaves these columns unblocked on purpose. BUD2026 re-applies
# the same three collection-blocking rules here so its own final file still
# excludes intercompany/blocked rows, regardless of what By_Customer sends.
ZERO_QUARTER_CUSTOMER_KEYWORDS = ("MINDWARE", "AKLANIAT", "IFIX")
ZERO_COLLECTION_MAIN_ACCOUNTS = {"12302", "12304", "12306"}
ALLOWED_COLLECTION_STATUSES = ("GOOD", "REGULAR", "SUBSTANDARD")

# ====================== MAIN MAPPER ======================

def map_by_customer_to_bud2026(
    df_customer: pd.DataFrame,
    ins_df: pd.DataFrame = None,
    selected_quarter: str = "Q1",
) -> pd.DataFrame:
    """
    Map input customer DataFrame to the BUD2026 quarterly model rows:
      - Identifiers
      - Insurance (from master; 0 when uninsured - the model's MIN(bucket, ins)
        chains need a number, never blank)
      - AR / Aging columns incl. the Not Due 0-30/.../180+ breakdown
      - AR Balance
      - Collections FC pre-fill for Q1-Q4 2026 (NaN when the By_Customer file
        has no source column for that quarter -> written blank)
    """
    work = df_customer.copy()
    out = pd.DataFrame(index=work.index)

    # ---------------- Identifiers ----------------
    out["CustCode"]            = _series_or_empty(work, "Cust Code").astype(str).str.strip()
    out["Cust Name"]           = _series_or_empty(work, "Cust Name").astype(str)
    out["BT"]                  = ""
    out["Sales Budget region"] = _derive_sales_budget_region(work).astype(str)
    out["Cust Region"]         = _series_or_empty(work, "Cust Region").astype(str)
    status_col = "Updated Status" if "Updated Status" in work.columns else "Customer Status"
    out["Customer Status"]     = _series_or_empty(work, status_col).astype(str)
    out["Main Ac"]             = normalize_excel_identifier_series(_series_or_empty(work, "Main Ac"))
    out["Focus List"]          = ""  # not exported; kept for the dashboard

    # ---------------- Insurance ----------------
    insurance = pd.Series([np.nan] * len(out), index=out.index)
    if ins_df is not None and not ins_df.empty:
        master = ins_df.copy()
        master["Customer Code"] = master.get("Customer Code", "").astype(str).str.strip()
        if "Main Account" in master.columns:
            master["Main Account"] = normalize_excel_identifier_series(master["Main Account"])
        else:
            master["Main Account"] = ""

        tmp = out[["CustCode", "Main Ac"]].copy()
        tmp["__CustCode"] = tmp["CustCode"].astype(str).str.strip()
        tmp["__MainAc"] = normalize_excel_identifier_series(tmp["Main Ac"])

        exact_master = master[master["Main Account"] != ""].copy()
        exact_match = pd.DataFrame(index=tmp.index)
        if not exact_master.empty:
            exact_match = tmp.merge(
                exact_master[["Customer Code", "Main Account", "Insurance Limit"]],
                how="left",
                left_on=["__CustCode", "__MainAc"],
                right_on=["Customer Code", "Main Account"],
            )

        if "Insurance Limit" in exact_match.columns:
            insurance = pd.to_numeric(exact_match["Insurance Limit"], errors="coerce")
            insurance.index = tmp.index

        needs_fallback = insurance.isna()
        if needs_fallback.any():
            fallback_master = master.drop_duplicates(subset=["Customer Code"], keep="first")
            fallback_match = tmp.loc[needs_fallback, ["__CustCode"]].merge(
                fallback_master[["Customer Code", "Insurance Limit"]],
                how="left",
                left_on="__CustCode",
                right_on="Customer Code",
            )
            fallback_values = pd.to_numeric(fallback_match["Insurance Limit"], errors="coerce")
            insurance.loc[needs_fallback] = fallback_values.values

    # 0 (never blank) for uninsured customers, like the master file
    out["Insurance"] = insurance.fillna(0.0)

    # ---------------- AR / Aging Columns ----------------
    on_acc_src   = _first_present(work, ["On Account (Derived)", "On account"])
    not_due_src  = _first_present(work, ["Not Due Amount", "Not Due (Derived)", "Not Due"])
    a1_30_src    = _first_present(work, ["Aging 1 to 30"])
    a31_60_src   = _first_present(work, ["Aging 31 to 60"])
    a61_90_src   = _first_present(work, ["Aging 61 to 90"])
    a91_120_src  = _first_present(work, ["Aging 91 to 120"])
    a121_150_src = _first_present(work, ["Aging 121 to 150"])
    a_ge_151_src = _first_present(work, ["Aging >=151", "Aging ≥151 (Amount)"])

    on_acc   = _num(work, on_acc_src)
    not_due  = _num(work, not_due_src)
    a1_30    = _num(work, a1_30_src)
    a31_60   = _num(work, a31_60_src)
    a61_90   = _num(work, a61_90_src)
    a91_120  = _num(work, a91_120_src)
    a121_150 = _num(work, a121_150_src)
    a_ge_151 = _num(work, a_ge_151_src)

    # AR Balance: use existing if available else sum separate aging buckets
    ar_balance_src = _first_present(work, ["AR Balance", "Ar Balance (Copy)", "Ar Balance"])
    if ar_balance_src:
        ar_bal = _num(work, ar_balance_src)
    else:
        ar_bal = on_acc + not_due + a1_30 + a31_60 + a61_90 + a91_120 + a121_150 + a_ge_151

    # ---------------- Not Due breakdown (K-O) ----------------
    breakdown_srcs = {name: _first_present(work, [src]) for name, src in NOT_DUE_BREAKDOWN_COLS.items()}
    used_breakdown = all(breakdown_srcs.values())
    if used_breakdown:
        for name, src in breakdown_srcs.items():
            out[name] = _num(work, src)
    else:
        # fallback: whole Not Due total into 'Not Due 0-30 days'
        out["Not Due\n0-30 days"] = not_due
        for name in list(NOT_DUE_BREAKDOWN_COLS)[1:]:
            out[name] = 0.0
    out.attrs["used_not_due_breakdown"] = used_breakdown

    # Re-apply collection blocking to the K-O breakdown (Orion's By_Customer
    # sheet intentionally leaves these unblocked - see NOT_DUE_BREAKDOWN_COLS
    # comment above).
    blocked_customer = out["Cust Name"].str.upper().str.contains(
        "|".join(ZERO_QUARTER_CUSTOMER_KEYWORDS), na=False
    )
    blocked_main_account = out["Main Ac"].isin(ZERO_COLLECTION_MAIN_ACCOUNTS)
    blocked_status = ~out["Customer Status"].str.upper().isin(ALLOWED_COLLECTION_STATUSES)
    blocked_row = blocked_customer | blocked_main_account | blocked_status
    for name in NOT_DUE_BREAKDOWN_COLS:
        out[name] = out[name].where(~blocked_row, 0.0)

    # ---------------- Quarter Collections FC pre-fill ----------------
    cfg = build_customer_output_config(selected_quarter)
    idx = QUARTER_ORDER.index(selected_quarter)
    collection_sources = {}
    for q_pos, quarter in enumerate(QUARTER_ORDER):
        if q_pos < idx:
            collection_sources[quarter] = None            # past quarter: leave blank
        elif q_pos == idx:
            collection_sources[quarter] = cfg["forecasted_label"]
        elif q_pos == idx + 1:
            collection_sources[quarter] = cfg["forecast_label"]
        else:
            collection_sources[quarter] = f"{quarter}-2026"
    for quarter, header in QUARTER_COLLECTION_HEADERS.items():
        src = collection_sources[quarter]
        src = _first_present(work, [src]) if src else None
        out[header] = pd.to_numeric(work[src], errors="coerce") if src else np.nan

    # ---------------- Map to BUD headers ----------------
    out["On\nAccount"]        = on_acc
    out["Not Due\nAmount"]    = not_due
    out["Aging\n1 to 30"]     = a1_30
    out["Aging\n31 to 60"]    = a31_60
    out["Aging\n61 to 90"]    = a61_90
    out["Aging\n91 to 120"]   = a91_120
    out["Aging\n121 to 150"]  = a121_150
    out["Aging\n>=151"]       = a_ge_151
    out[" AR\nBalance"]       = ar_bal

    return out
