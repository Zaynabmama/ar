"""Cleaning rules + derived columns, replicating the manual process exactly.

Every rule here was confirmed against the real files and the real formulas
in the existing manually-built workbook (not guessed) -- see the helper
column formulas in 'AR Provision New' columns BJ:BY of the sample output,
and the pivotTable/pivotCache XML for how they feed the four pivots.
"""

from __future__ import annotations

import pandas as pd

from ar_report.lookups import Lookups, bucket_label, resolve_region

DROPPED_MAIN_ACCOUNTS = {12302, 12304, 12306}
AGING_CATEGORIES = [
    "On account",
    "Not Due",
    "Aging 1 to 30",
    "Aging 31 to 60",
    "Aging 61 to 90",
    "Aging 91 to 120",
    "Aging 121 to 150",
    "Aging >=151",
]


def resolve_as_on_date(ar_df: pd.DataFrame, override=None) -> pd.Timestamp:
    if override is not None:
        return pd.Timestamp(override)
    if "As Of Dt" in ar_df.columns:
        values = pd.to_datetime(ar_df["As Of Dt"], errors="coerce").dropna()
        if len(values):
            return values.mode().iloc[0]
    raise ValueError(
        "AR Provision export has no usable 'As Of Dt' value -- pass an explicit as_on_date."
    )


def clean_ar_provision(df: pd.DataFrame, lookups) -> pd.DataFrame:
    """Applies the three manual filters: intercompany-entity exclusions,
    main-account exclusions, and the near-zero AR balance noise filter.

    NOTE: "remove the mindware entities (aklaniat & IFIX)" from the original
    process description does NOT mean drop every row where Legal Entity
    contains "Aklaniat" -- verified against a real, hand-built report: of
    2,129 Aklaniat-legal-entity customers in a raw export, 1,100 are
    legitimately kept (Legal Entity there just means which Mindware book-entity
    booked the AR, unrelated to the customer). The real rule is a specific,
    maintained list of customer codes that represent *other Mindware group
    entities themselves* (intercompany accounts, not real customers) --
    "aklaniat" and "IFIX" are two examples from that list, not the whole
    rule. See template/lookup_data.xlsx -> IntercompanyEntities.
    """
    # The raw export carries a trailing "Summary:" row (Cust Code literally
    # "Summary:", every other field blank) -- a report-footer artifact, not
    # a real record. Nothing else here would catch it (Ar Balance/Main Ac
    # are NaN, not in-range or in the dropped-account set), so it would
    # otherwise leak into By Customer/By Invoice as a garbage row. A real
    # AR line always has a Document Number; this is the same fix pattern
    # as the backlog report's stray subtotal row.
    out = df.dropna(subset=["Document Number"]).copy()

    cust_code = out["Cust Code"].astype("string").str.strip()
    is_intercompany = cust_code.isin(lookups.intercompany_cust_codes)

    main_ac = pd.to_numeric(out["Main Ac"], errors="coerce")
    is_dropped_account = main_ac.isin(DROPPED_MAIN_ACCOUNTS)

    ar_balance = pd.to_numeric(out["Ar Balance"], errors="coerce")
    is_near_zero = ar_balance.between(-5, 5)

    keep = ~(is_intercompany | is_dropped_account | is_near_zero)
    return out.loc[keep].reset_index(drop=True)


def clean_backlog(df: pd.DataFrame) -> pd.DataFrame:
    """Drops orders with Credit Status 'NA', plus rows with no Order number.

    The raw export carries a stray row with every column blank except
    Pending Val (Lc) -- a subtotal/grand-total artifact rather than a real
    order (confirmed against a real report: this single row alone accounted
    for a ~$763M overstatement, dwarfing everything else). A real backlog
    entry always has an Order number, so this is a safe, general filter
    rather than a one-off patch for this specific stray value.
    """
    out = df.dropna(subset=["Order"])
    status = out["Credit Status"].astype("string").fillna("").str.strip()
    keep = status.str.casefold() != "na"
    return out.loc[keep].reset_index(drop=True)


def add_helper_columns(ar_df: pd.DataFrame, as_on_date: pd.Timestamp, lookups: Lookups) -> pd.DataFrame:
    """Reproduces the 16 'blue' helper columns (BJ:BY in the original sheet)
    that the four pivots and the By Invoice sheet are built from."""
    out = ar_df.copy()
    as_on_date = pd.Timestamp(as_on_date).normalize()
    month_end = as_on_date + pd.offsets.MonthEnd(0)

    doc_date = pd.to_datetime(out["Document Date"], errors="coerce")
    due_date = pd.to_datetime(out["Document Due Date"], errors="coerce")
    ar_balance = pd.to_numeric(out["Ar Balance"], errors="coerce").fillna(0)

    out["Ageing"] = (as_on_date - doc_date).dt.days
    out["Overdue days"] = (as_on_date - due_date).dt.days

    out["Region"] = [
        resolve_region(code, region_raw, lookups)
        for code, region_raw in zip(out["Cust Code"], out["Cust Region"])
    ]

    # Uploaded city/sub-region overrides (e.g. AUH) show up in the Country
    # field itself, not the broader Region classification above -- Region
    # was already computed from the original country a moment ago, so this
    # doesn't affect it either way.
    cust_code_stripped = out["Cust Code"].astype("string").str.strip()
    country_override = cust_code_stripped.map(lookups.reporting_region_override)
    override_mask = country_override.notna()
    out.loc[override_mask, "Cust Region"] = country_override[override_mask]

    out["Ar Balance2"] = ar_balance

    out["Aging Bracket"] = [
        "On account" if bal < 0 else bucket_label(days, lookups.ageing_brackets)
        for bal, days in zip(ar_balance, out["Overdue days"])
    ]

    customer_status = out["Customer Status"].astype("string").fillna("").str.strip()
    out["Updated Status"] = customer_status.where(customer_status != "", "SUBSTANDARD")

    out["Additional due End of month"] = ar_balance.where(
        (due_date > as_on_date) & (due_date <= month_end), 0
    )

    out["Invoice Value"] = ar_balance.where(ar_balance > 0, 0)
    out["On Account2"] = pd.to_numeric(out["On Account"], errors="coerce").fillna(0)
    out["Not Due (helper)"] = pd.to_numeric(out["Not Due Amount"], errors="coerce").fillna(0)

    # Cascading aging buckets: each bucket = amount-if-over-threshold minus
    # everything already claimed by a higher bucket (mirrors the BT:BY
    # formulas exactly, thresholds pulled from the ageing brackets table).
    thresholds = [low - 1 for low, _high, label in lookups.ageing_brackets if label != "Not Due"]
    bucket_cols = [
        "Aging 1 to 30 (calc)",
        "Aging 31 to 60 (calc)",
        "Aging 61 to 90 (calc)",
        "Aging 91 to 120 (calc)",
        "Aging 121 to 150 (calc)",
        "Aging >=151 (calc)",
    ]
    overdue_days = out["Overdue days"].fillna(-1)
    already_claimed = pd.Series(0.0, index=out.index)
    computed = {}
    for threshold, col in zip(reversed(thresholds), reversed(bucket_cols)):
        bucket_amount = out["Invoice Value"].where(overdue_days > threshold, 0) - already_claimed
        computed[col] = bucket_amount
        already_claimed = already_claimed + bucket_amount
    for col in bucket_cols:
        out[col] = computed[col]

    return out
