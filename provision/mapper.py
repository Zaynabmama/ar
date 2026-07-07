from datetime import date

import numpy as np
import pandas as pd

from common.identifier_utils import normalize_excel_identifier_series
from orion.processor import sanitize_colnames

# Sheets of the tool-1 output workbook that hold invoice-level rows, in
# preference order. Used to compute the Not Due breakdown (columns L-P).
INVOICE_SHEET_CANDIDATES = ["Invoice", "AR_Backlog", "Traverse_AR"]


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


def build_not_due_breakdown(invoice_df: pd.DataFrame | None, ar_date: date) -> pd.DataFrame | None:
    """Compute the Not Due breakdown (columns L-P) per customer from an
    invoice-level sheet (Invoice / AR_Backlog / Traverse_AR).

    Buckets by Document Due Date relative to the AR Data Date, mirroring the
    original model's SUMIFS conditions:
      L: due in (ar, ar+30]     M: (ar+30, ar+60]     N: (ar+60, ar+90]
      O: (ar+90, ar+180]        P: (ar+180, 31-Dec-2026]  (2027+ excluded)
    """
    if invoice_df is None or invoice_df.empty:
        return None

    work = sanitize_colnames(invoice_df.copy())
    work = work.loc[:, ~work.columns.duplicated(keep="first")]

    code_col = _first_present(work, ["Cust Code", "CustId"])
    due_col = _first_present(work, ["Document Due Date", "GrpDue"])
    amount_col = _first_present(work, ["Ar Balance", "invoice value - USD", "Invoice Value (Derived)"])
    if not code_col or not due_col or not amount_col:
        return None

    inv = pd.DataFrame(index=work.index)
    inv["code"] = work[code_col].astype(str).str.strip()
    inv["due"] = pd.to_datetime(work[due_col], errors="coerce")
    inv["amount"] = pd.to_numeric(work[amount_col], errors="coerce").fillna(0.0)
    inv = inv[(inv["amount"] > 0) & inv["due"].notna() & (inv["code"] != "")]
    if inv.empty:
        return None

    ar_ts = pd.Timestamp(ar_date)
    days = (inv["due"] - ar_ts).dt.days
    year_end = pd.Timestamp(2026, 12, 31)
    buckets = {
        "L": (days > 0) & (days <= 30),
        "M": (days > 30) & (days <= 60),
        "N": (days > 60) & (days <= 90),
        "O": (days > 90) & (days <= 180),
        "P": (days > 180) & (inv["due"] <= year_end),
    }
    out = pd.DataFrame({"code": inv["code"]})
    for col, mask in buckets.items():
        out[col] = inv["amount"].where(mask, 0.0)
    return out.groupby("code", as_index=True).sum(numeric_only=True)


def find_invoice_sheet(sheet_names: list[str]) -> str | None:
    for name in INVOICE_SHEET_CANDIDATES:
        if name in sheet_names:
            return name
    return None


def map_by_customer_to_provision(
    df_customer: pd.DataFrame,
    ins_df: pd.DataFrame | None = None,
    invoice_df: pd.DataFrame | None = None,
    ar_date: date | None = None,
) -> tuple[pd.DataFrame, bool]:
    """Map the By_Customer sheet to the fixed columns (A-Y) of the provision
    forecast 'ALL' sheet. Returns (df_fixed keyed by column letter, used_invoice).

    The Not Due breakdown (L-P) is computed from the invoice-level sheet when
    available; otherwise the whole Not Due total goes to column L (used_invoice
    is False in that case).
    """
    work = sanitize_colnames(df_customer.copy())
    work = work.loc[:, ~work.columns.duplicated(keep="last")]

    out = pd.DataFrame(index=work.index)
    out["A"] = _series_or_empty(work, "Cust Code").astype(str).str.strip()   # CustCode
    out["B"] = _series_or_empty(work, "Cust Name").fillna("").astype(str)    # Cust Name
    out["C"] = ""                                                            # BT
    out["D"] = _series_or_empty(work, "Cust Region").fillna("").astype(str)  # Country
    region_col = "Region" if "Region" in work.columns else "Cust Region"
    out["E"] = _series_or_empty(work, region_col).fillna("").astype(str)     # Cust Region
    status_col = "Updated Status" if "Updated Status" in work.columns else "Customer Status"
    out["F"] = _series_or_empty(work, status_col).fillna("").astype(str)     # Customer Status
    out["G"] = normalize_excel_identifier_series(_series_or_empty(work, "Main Ac"))  # Main Ac
    out["H"] = ""                                                            # Comments
    out["I"] = lookup_insurance(out["A"], out["G"], ins_df)                  # Insurance

    on_acc_col = _first_present(work, ["On account", "On Account (Derived)"])
    not_due_col = _first_present(work, ["Not Due", "Not Due Amount"])
    out["J"] = _num(work, on_acc_col)                                        # On Account
    out["K"] = _num(work, not_due_col)                                       # Not Due Amount
    out["Q"] = _num(work, _first_present(work, ["Aging 1 to 30"]))
    out["R"] = _num(work, _first_present(work, ["Aging 31 to 60"]))
    out["S"] = _num(work, _first_present(work, ["Aging 61 to 90"]))
    out["T"] = _num(work, _first_present(work, ["Aging 91 to 120"]))
    out["U"] = _num(work, _first_present(work, ["Aging 121 to 150"]))
    out["V"] = _num(work, _first_present(work, ["Aging >=151", "Aging ≥151"]))
    ar_balance_col = _first_present(work, ["Ar Balance", "AR Balance", "Ar Balance (Copy)"])
    if ar_balance_col:
        out["W"] = _num(work, ar_balance_col)                                # AR Balance
    else:
        out["W"] = out[["J", "K", "Q", "R", "S", "T", "U", "V"]].sum(axis=1)
    out["X"] = np.nan                                                        # AR Provision at <prior> - manual
    out["Y"] = np.nan                                                        # AR Provision at 31-12-2025 - manual

    out = out[out["A"].str.strip().ne("") & out["A"].str.lower().ne("nan")]

    breakdown = build_not_due_breakdown(invoice_df, ar_date) if ar_date else None
    if breakdown is not None:
        for col in ["L", "M", "N", "O", "P"]:
            out[col] = out["A"].map(breakdown[col]).fillna(0.0)
        used_invoice = True
    else:
        # fallback: whole Not Due total into 'Not Due 0-30 days' (column L)
        out["L"] = out["K"]
        out["M"] = 0.0
        out["N"] = 0.0
        out["O"] = 0.0
        out["P"] = 0.0
        used_invoice = False

    out = out.sort_values("W", ascending=False).reset_index(drop=True)
    return out, used_invoice
