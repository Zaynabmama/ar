"""Pandas equivalents of the workbook's four Excel PivotTables, plus the two
downstream 'By Customer' / 'By Invoice' views that read from them.

Field roles (row/column/data fields, and which source column feeds each)
were read directly out of xl/pivotTables/pivotTable*.xml and
xl/pivotCache/pivotCacheDefinition*.xml in the sample workbook, not guessed.
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from ar_report.transform import AGING_CATEGORIES

CASCADE_COLS = [
    "Aging 1 to 30 (calc)",
    "Aging 31 to 60 (calc)",
    "Aging 61 to 90 (calc)",
    "Aging 91 to 120 (calc)",
    "Aging 121 to 150 (calc)",
    "Aging >=151 (calc)",
]


def pivot_customer_aging(ar_df: pd.DataFrame) -> pd.DataFrame:
    """Sheet 'Pivot', table 1: Sum of Ar Balance2 by customer x aging bracket."""
    group_cols = ["Cust Code", "Cust Name", "Main Ac", "Cust Region", "Region", "Updated Status"]
    pt = ar_df.pivot_table(
        index=group_cols,
        columns="Aging Bracket",
        values="Ar Balance2",
        aggfunc="sum",
        fill_value=0,
    )
    for cat in AGING_CATEGORIES:
        if cat not in pt.columns:
            pt[cat] = 0
    pt = pt[AGING_CATEGORIES]
    pt["Grand Total"] = pt.sum(axis=1)
    return pt.reset_index().sort_values(["Cust Code", "Main Ac"]).reset_index(drop=True)


def pivot_additional_dues(ar_df: pd.DataFrame) -> pd.DataFrame:
    """Sheet 'Pivot', table 2: Sum of Additional due End of month by customer."""
    return (
        ar_df.groupby("Cust Code", as_index=False)["Additional due End of month"]
        .sum()
        .rename(columns={"Additional due End of month": "Sum of Additional due End of month"})
        .sort_values("Cust Code")
        .reset_index(drop=True)
    )


def pivot_backlog_detail(backlog_df: pd.DataFrame) -> pd.DataFrame:
    """Sheet 'Backlog' embedded pivot: rolls up identical
    (Order Date, Customer, Order, LPO, End User, Product Line, Term) combos."""
    group_cols = [
        "Order Date",
        "Customer Code",
        "Customer Name",
        "Order",
        "Customer LPO",
        "End User Name",
        "Product Line Desc",
        "Term Code",
    ]
    return (
        backlog_df.groupby(group_cols, as_index=False, dropna=False)["Pending Val (Lc)"]
        .sum()
        .rename(columns={"Pending Val (Lc)": "Sum of Pending Val (Lc)"})
        .sort_values("Order Date")
        .reset_index(drop=True)
    )


def pivot_backlog_by_customer(backlog_df: pd.DataFrame) -> pd.DataFrame:
    """'Look up links' PIVOT for Backlog: Sum of Pending Val (Lc) by customer."""
    return (
        backlog_df.groupby("Customer Code", as_index=False)["Pending Val (Lc)"]
        .sum()
        .rename(columns={"Customer Code": "Row Labels", "Pending Val (Lc)": "Sum of Pending Val (Lc)"})
        .sort_values("Row Labels")
        .reset_index(drop=True)
    )


def build_by_customer(
    customer_pivot: pd.DataFrame,
    additional_dues: pd.DataFrame,
    backlog_by_customer: pd.DataFrame,
    insurance_df: pd.DataFrame,
    lookups,
) -> pd.DataFrame:
    """Replicates the VLOOKUP/XLOOKUP chain in 'By Customer'. Where the
    original used VLOOKUP against the customer pivot (which is keyed by
    Cust Code + Main Ac), it silently takes the FIRST matching row for a
    customer with more than one Main Ac -- kept as-is here for parity with
    the numbers the team already sees; flagged separately as worth revisiting."""
    first_per_customer = customer_pivot.drop_duplicates(subset="Cust Code", keep="first").set_index(
        "Cust Code"
    )
    additional_dues_by_code = additional_dues.set_index("Cust Code")["Sum of Additional due End of month"]
    backlog_by_code = backlog_by_customer.set_index("Row Labels")["Sum of Pending Val (Lc)"]

    insurance_by_code = insurance_df.drop_duplicates(subset="Customer Code").set_index("Customer Code")

    rows = []
    for cust_code, sub in first_per_customer.iterrows():
        cust_name = sub["Cust Name"]
        on_account = sub.get("On account", 0)
        not_due = sub.get("Not Due", 0)
        aging_1_30 = sub.get("Aging 1 to 30", 0)
        aging_31_60 = sub.get("Aging 31 to 60", 0)
        aging_61_90 = sub.get("Aging 61 to 90", 0)
        aging_91_120 = sub.get("Aging 91 to 120", 0)
        aging_121_150 = sub.get("Aging 121 to 150", 0)
        aging_151 = sub.get("Aging >=151", 0)
        overdue = aging_1_30 + aging_31_60 + aging_61_90 + aging_91_120 + aging_121_150 + aging_151
        ar_balance = overdue + not_due + on_account
        overdue_plus_on_account = overdue + on_account
        remaining_month_end = additional_dues_by_code.get(cust_code, 0)
        total_overdue_till_month_end = (
            overdue_plus_on_account + remaining_month_end
            if (overdue_plus_on_account + remaining_month_end) > 0
            else 0
        )

        ct_dso = lookups.credit_terms_by_customer_name.get(str(cust_name).strip(), (0, 0))
        ins_row = insurance_by_code.loc[cust_code] if cust_code in insurance_by_code.index else None

        rows.append(
            {
                "Cust Code": cust_code,
                "Cust Name": cust_name,
                "Main Ac": sub["Main Ac"],
                "EZ#": ins_row["Buyer Insurance Code"] if ins_row is not None else "",
                "Country": sub["Cust Region"],
                "Region": sub["Region"],
                "CustStatus": sub["Updated Status"] or "SUBSTANDARD",
                "CT": ct_dso[0],
                "DSO": ct_dso[1],
                "Insured Limit": ins_row["Insurance Limit"] if ins_row is not None else 0,
                "Backlog": backlog_by_code.get(cust_code, 0),
                "On account": on_account,
                "Not Due": not_due,
                "Ar Balance": ar_balance,
                "Overdue + On Account": overdue_plus_on_account,
                "Overdue": overdue,
                "Aging 1 to 30": aging_1_30,
                "Aging 31 to 60": aging_31_60,
                "Aging 61 to 90": aging_61_90,
                "Aging 91 to 120": aging_91_120,
                "Aging 121 to 150": aging_121_150,
                "Aging >=151": aging_151,
                "Remaining Month End Dues": remaining_month_end,
                "TOTAL  Overdues till Month End": total_overdue_till_month_end,
            }
        )
    return pd.DataFrame(rows)


def build_by_invoice(ar_df: pd.DataFrame) -> pd.DataFrame:
    """'By Invoice': a reshaped/renamed copy of the cleaned AR data, one row
    per document, plus the two age-in-days columns."""
    out = pd.DataFrame(
        {
            "Cust Code": ar_df["Cust Code"],
            "Cust Name": ar_df["Cust Name"],
            "Main Ac": ar_df["Main Ac"],
            "Cust Region": ar_df["Cust Region"],
            "Document Number": ar_df["Document Number"],
            "Document Date": ar_df["Document Date"],
            "Document Due Date": ar_df["Document Due Date"],
            "Ageing": ar_df["Ageing"],
            "Overdue days": ar_df["Overdue days"],
            "Payment Terms": ar_df["Payment Terms"],
            "On Account": ar_df["On Account"],
            "Not Due": ar_df["Not Due Amount"],
            "Ar Balance": ar_df["Ar Balance"],
            "Aging 1 to 30": ar_df["Aging 1 to 30 (calc)"],
            "Aging 31 to 60": ar_df["Aging 31 to 60 (calc)"],
            "Aging 61 to 90": ar_df["Aging 61 to 90 (calc)"],
            "Aging 91 to 120": ar_df["Aging 91 to 120 (calc)"],
            "Aging 121 to 150": ar_df["Aging 121 to 150 (calc)"],
            "Aging >=151": ar_df["Aging >=151 (calc)"],
            "Brand": ar_df["Brand"],
            "Total Insurance Limit": ar_df["Total Insurance Limit"],
            "LC & BG Guarantee": ar_df["LC & BG Guarantee"],
            "SO No": ar_df["SO No"],
            "LPO No": ar_df["LPO No"],
        }
    )
    return out.reset_index(drop=True)
