"""Static reference tables the daily data gets joined against.

These are NOT part of the daily AR/PDC/Backlog/Insurance exports -- they're
maintained by hand (region mapping, per-customer credit-terms/DSO history,
and manual region overrides for a handful of customers). They live in
template/lookup_data.xlsx. To update them, edit that file directly in Excel;
the tool re-reads it on every run.
"""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd


@dataclass
class Lookups:
    region_by_country: dict
    ageing_brackets: list  # list of (low, high, label), ascending by low
    reporting_region_override: dict  # Cust Code -> Region
    credit_terms_by_customer_name: dict  # Cust Name -> (Credit Terms, DSO)
    intercompany_cust_codes: set  # Cust Codes representing other Mindware group entities

    # Raw tables, kept alongside the derived dicts above so build_output.py can
    # embed them in the output workbook for formulas to reference directly --
    # the dicts above are for Python-side computation (row lists, testing),
    # these are for writing real lookup tables into the "Look up links" sheet.
    region_df: pd.DataFrame
    ageing_df: pd.DataFrame
    trading_df: pd.DataFrame
    overrides_df: pd.DataFrame


def load_lookups(template_path) -> Lookups:
    region_df = pd.read_excel(template_path, sheet_name="Region", engine="openpyxl")
    ageing_df = pd.read_excel(template_path, sheet_name="AgeingBrackets", engine="openpyxl")
    trading_df = pd.read_excel(template_path, sheet_name="TradingExperience", engine="openpyxl")
    overrides_df = pd.read_excel(template_path, sheet_name="ReportingRegionOverrides", engine="openpyxl")
    intercompany_df = pd.read_excel(template_path, sheet_name="IntercompanyEntities", engine="openpyxl")

    region_by_country = {
        str(row["Country"]).strip().casefold(): row["Region"]
        for _, row in region_df.dropna(subset=["Country"]).iterrows()
    }

    ageing_brackets = sorted(
        (
            (float(row["Low"]), float(row["High"]), str(row["Label"]))
            for _, row in ageing_df.dropna(subset=["Low", "High", "Label"]).iterrows()
        ),
        key=lambda t: t[0],
    )

    reporting_region_override = {
        str(row["Cust Code"]).strip(): row["Reporting Region"]
        for _, row in overrides_df.dropna(subset=["Cust Code"]).iterrows()
    }

    credit_terms_by_customer_name = {}
    for _, row in trading_df.dropna(subset=["Customer Name"]).iterrows():
        name = str(row["Customer Name"]).strip()
        ct = row.get("Credit Terms")
        dso = row.get("DSO")
        credit_terms_by_customer_name[name] = (
            0 if pd.isna(ct) else ct,
            0 if pd.isna(dso) else dso,
        )

    intercompany_cust_codes = {
        str(code).strip() for code in intercompany_df["Cust Code"].dropna()
    }

    return Lookups(
        region_by_country=region_by_country,
        ageing_brackets=ageing_brackets,
        reporting_region_override=reporting_region_override,
        credit_terms_by_customer_name=credit_terms_by_customer_name,
        intercompany_cust_codes=intercompany_cust_codes,
        region_df=region_df,
        ageing_df=ageing_df,
        trading_df=trading_df,
        overrides_df=overrides_df,
    )


def bucket_label(overdue_days, ageing_brackets) -> str:
    """Replicates the approximate-match VLOOKUP against the ageing brackets table."""
    if pd.isna(overdue_days):
        overdue_days = 0
    label = ageing_brackets[0][2]
    for low, _high, lbl in ageing_brackets:
        if overdue_days >= low:
            label = lbl
        else:
            break
    return label


def apply_region_overrides(lookups: Lookups, override_dfs: list[pd.DataFrame]) -> None:
    """Merges uploaded city/sub-region customer lists (e.g. "AUH Customers
    list") into the region-override map, in place. Later files win on a
    Cust Code collision; uploaded entries take priority over whatever was
    already in the static template, since these are the fresher, per-run
    source of truth the team maintains by uploading a current list."""
    for df in override_dfs:
        for _, row in df.dropna(subset=["Cust Code"]).iterrows():
            code = str(row["Cust Code"]).strip()
            region = row["Reporting Region"]
            if pd.notna(region) and str(region).strip():
                lookups.reporting_region_override[code] = str(region).strip()


def resolve_region(cust_code, cust_region_raw, lookups: Lookups):
    """The broad Region classification (GCC/QNAL/KSA) -- always computed
    from the country, never from the city/sub-region overrides (those land
    in the Country field itself instead; see apply_region_overrides and
    where 'Cust Region' gets overwritten in transform.add_helper_columns)."""
    code = str(cust_code).strip()
    if code.upper().startswith("CK"):
        return "KSA"
    return lookups.region_by_country.get(str(cust_region_raw).strip().casefold(), "")
