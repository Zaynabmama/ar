import re

import pandas as pd

from common.region_maps import classify_region
from traverse.rules import (
    MAIN_ACCOUNT_BY_CUSTOMER_CODE,
    REGION_BY_COUNTRY,
    STATUS_BY_CUSTOMER_CODE,
    lookup_with_default,
    normalize_key,
)


def sanitize_colnames(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = []
    for col in df.columns:
        name = str(col).replace("\u00A0", " ").strip()
        name = re.sub(r"\s+", " ", name)
        cleaned.append(name)

    out = df.copy()
    out.columns = cleaned
    return out


def _series_or_blank(df: pd.DataFrame, col: str) -> pd.Series:
    if col in df.columns:
        subset = df.loc[:, col]
        if isinstance(subset, pd.DataFrame):
            return subset.iloc[:, 0]
        return subset
    return pd.Series([""] * len(df), index=df.index)


def _normalise_duplicate_headers(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [re.sub(r"\.\d+$", "", str(col)) for col in out.columns]
    return out


def _score_traverse_sheet(df: pd.DataFrame) -> int:
    expected = {"CustId", "CustName", "GrossAmount", "GrossAmountBase", "GrpDate", "GrpDue", "Country"}
    cols = {str(c).replace("\u00A0", " ").strip() for c in df.columns}
    return len(expected & cols) + (1 if len(df) > 0 else 0)


def _read_traverse_sheet(file_obj) -> pd.DataFrame:
    xl = pd.ExcelFile(file_obj, engine="openpyxl")
    best_df = None
    best_score = -1

    for sheet_name in xl.sheet_names:
        try:
            df = pd.read_excel(xl, sheet_name=sheet_name, dtype=str)
        except Exception:
            continue
        df = sanitize_colnames(df)
        score = _score_traverse_sheet(df)
        if score > best_score:
            best_df = df
            best_score = score

    if best_df is None:
        raise ValueError("Could not read any sheet from the Traverse workbook.")
    return best_df


def prepare_traverse_input(file_obj) -> pd.DataFrame:
    df = _read_traverse_sheet(file_obj)
    df = _normalise_duplicate_headers(df)
    return df


def _coface_limit_lookup_from_master(master_df: pd.DataFrame) -> dict[str, float]:
    lookup: dict[str, float] = {}
    if master_df is None or master_df.empty:
        return lookup

    master = master_df.copy()
    if "Customer reference" not in master.columns or "Amount agreed" not in master.columns:
        return lookup

    master["__ref"] = master["Customer reference"].map(normalize_key)
    master["__amount"] = pd.to_numeric(master["Amount agreed"], errors="coerce").fillna(0.0)
    master = master.dropna(subset=["__ref"])
    master = master[master["__ref"] != ""]
    master = master.drop_duplicates(subset=["__ref"], keep="last")

    for _, row in master.iterrows():
        lookup[str(row["__ref"])] = float(row["__amount"] or 0)

    return lookup


def enrich_traverse_lookups(df: pd.DataFrame, insurance_master_df: pd.DataFrame | None = None) -> pd.DataFrame:
    work = df.copy()

    cust_code = _series_or_blank(work, "CustId")
    country = _series_or_blank(work, "Country")
    cust_name = _series_or_blank(work, "CustName")

    work["Status"] = [lookup_with_default(STATUS_BY_CUSTOMER_CODE, key, "Regular") for key in cust_code]
    work["Customer Region"] = [lookup_with_default(REGION_BY_COUNTRY, key, "") for key in country]
    work["Main Acccount"] = [lookup_with_default(MAIN_ACCOUNT_BY_CUSTOMER_CODE, key, "12301") for key in cust_code]
    coface_lookup = _coface_limit_lookup_from_master(insurance_master_df)
    if coface_lookup:
        work["Credit Limit Coface"] = [
            float(lookup_with_default(coface_lookup, key, 0) or 0) for key in cust_code
        ]
    else:
        work["Credit Limit Coface"] = [0.0 for _ in cust_name]

    return work


def prepare_traverse_output(df: pd.DataFrame, insurance_master_df: pd.DataFrame | None = None) -> pd.DataFrame:
    return enrich_traverse_lookups(df, insurance_master_df=insurance_master_df)
