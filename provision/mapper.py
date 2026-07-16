import numpy as np
import pandas as pd

from common.identifier_utils import normalize_excel_identifier_series
from orion.processor import sanitize_colnames


def _series_or_empty(df: pd.DataFrame, col: str) -> pd.Series:
    if col in df.columns:
        return df[col]
    return pd.Series([""] * len(df), index=df.index)


def _num(df: pd.DataFrame, col: str) -> pd.Series:
    if not col or col not in df.columns:
        return pd.Series([0.0] * len(df), index=df.index, dtype="float64")
    return pd.to_numeric(df[col], errors="coerce").fillna(0.0)


def _first_present(df: pd.DataFrame, candidates: list[str]) -> str | None:
    normalized = {"".join(str(c).split()).lower(): c for c in df.columns}
    for c in candidates:
        hit = normalized.get("".join(c.split()).lower())
        if hit is not None:
            return hit
    return None


def lookup_insurance(cust_codes: pd.Series, main_acs: pd.Series, ins_df: pd.DataFrame | None) -> pd.Series:
    """Insurance Limit per (Customer Code, Main Account), falling back to Customer Code only.

    Same behaviour as the BUD2026 mapper insurance block.
    """
    insurance = pd.Series([np.nan] * len(cust_codes), index=cust_codes.index)
    if ins_df is None or ins_df.empty:
        return insurance

    master = ins_df.copy()
    master["Customer Code"] = master.get("Customer Code", "").astype(str).str.strip()
    if "Main Account" in master.columns:
        master["Main Account"] = normalize_excel_identifier_series(master["Main Account"])
    else:
        master["Main Account"] = ""

    tmp = pd.DataFrame(index=cust_codes.index)
    tmp["__CustCode"] = cust_codes.astype(str).str.strip()
    tmp["__MainAc"] = normalize_excel_identifier_series(main_acs)

    exact_master = master[master["Main Account"] != ""]
    if not exact_master.empty:
        exact_match = tmp.merge(
            exact_master[["Customer Code", "Main Account", "Insurance Limit"]],
            how="left",
            left_on=["__CustCode", "__MainAc"],
            right_on=["Customer Code", "Main Account"],
        )
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
        insurance.loc[needs_fallback] = pd.to_numeric(
            fallback_match["Insurance Limit"], errors="coerce"
        ).values

    return insurance


# By_Customer columns feeding the model's K-O Not Due breakdown (written by
# tool 1 with embedded newlines; _first_present ignores all whitespace)
NOT_DUE_BREAKDOWN_COLS = {
    "K": "Not Due 0-30 days",
    "L": "Not Due 31-60 days",
    "M": "Not Due 61-90 days",
    "N": "Not Due 91-180 days",
    "O": "Not Due 180+ days",
}


def map_by_customer_to_provision(
    df_customer: pd.DataFrame,
    ins_df: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, bool]:
    """Map the By_Customer sheet to the fixed columns (A-U) of the provision
    forecast 'ALL' sheet (new Master File layout). Returns (df_fixed keyed by
    column letter, used_breakdown).

    The Not Due breakdown (K-O) is read from the By_Customer "Not Due ..."
    columns (collectible view, added by the AR Backlog tool); when they are
    missing the whole Not Due total goes to column K (used_breakdown is False).
    The AR Balance (V), prior provisions (W/X) and Notes (AC) are not emitted:
    V is a live formula and the rest are manual.
    """
    work = sanitize_colnames(df_customer.copy())
    work = work.loc[:, ~work.columns.duplicated(keep="last")]

    out = pd.DataFrame(index=work.index)
    out["A"] = _series_or_empty(work, "Cust Code").astype(str).str.strip()   # CustCode
    out["B"] = _series_or_empty(work, "Cust Name").fillna("").astype(str)    # Cust Name
    out["D"] = _series_or_empty(work, "Cust Region").fillna("").astype(str)  # Country
    region_col = "Region" if "Region" in work.columns else "Cust Region"
    out["E"] = _series_or_empty(work, region_col).fillna("").astype(str)     # Cust Region
    status_col = "Updated Status" if "Updated Status" in work.columns else "Customer Status"
    out["F"] = _series_or_empty(work, status_col).fillna("").astype(str)     # Customer Status
    out["G"] = normalize_excel_identifier_series(_series_or_empty(work, "Main Ac"))  # Main Ac
    # Insurance must be 0 (never blank) for uninsured customers: the model's
    # MIN(bucket, ins) chains ignore blank cells, silently insuring the oldest
    # bucket at the 5% rate. The master file stores 0 for all uninsured rows.
    out["H"] = lookup_insurance(out["A"], out["G"], ins_df).fillna(0.0)      # Insurance

    on_acc_col = _first_present(work, ["On account", "On Account (Derived)"])
    not_due_col = _first_present(work, ["Not Due", "Not Due Amount"])
    out["I"] = _num(work, on_acc_col)                                        # On Account
    out["J"] = _num(work, not_due_col)                                       # Not Due Amount
    out["P"] = _num(work, _first_present(work, ["Aging 1 to 30"]))
    out["Q"] = _num(work, _first_present(work, ["Aging 31 to 60"]))
    out["R"] = _num(work, _first_present(work, ["Aging 61 to 90"]))
    out["S"] = _num(work, _first_present(work, ["Aging 91 to 120"]))
    out["T"] = _num(work, _first_present(work, ["Aging 121 to 150"]))
    out["U"] = _num(work, _first_present(work, ["Aging >=151", "Aging ≥151"]))

    source_cols = {k: _first_present(work, [name]) for k, name in NOT_DUE_BREAKDOWN_COLS.items()}
    if all(source_cols.values()):
        for col, src in source_cols.items():
            out[col] = _num(work, src)
        used_breakdown = True
    else:
        # fallback: whole Not Due total into 'Not Due 0-30 days' (column K)
        out["K"] = out["J"]
        out["L"] = 0.0
        out["M"] = 0.0
        out["N"] = 0.0
        out["O"] = 0.0
        used_breakdown = False

    out = out[out["A"].str.strip().ne("") & out["A"].str.lower().ne("nan")]

    # sort by AR balance desc (column V in the model = I+J+P..U, written as a formula)
    ar_balance_col = _first_present(work, ["Ar Balance", "AR Balance", "Ar Balance (Copy)"])
    if ar_balance_col:
        sort_key = _num(work, ar_balance_col).reindex(out.index)
    else:
        sort_key = out[["I", "J", "P", "Q", "R", "S", "T", "U"]].sum(axis=1)
    out = out.loc[sort_key.sort_values(ascending=False).index].reset_index(drop=True)
    return out, used_breakdown
