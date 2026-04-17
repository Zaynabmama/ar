from datetime import datetime
from typing import Any, Dict, Iterable, Mapping, Tuple

QUARTER_ORDER = ["Q1", "Q2", "Q3", "Q4"]

QUARTER_END_LABELS_2026 = {
    "Q1": "31/03/2026",
    "Q2": "30/06/2026",
    "Q3": "30/09/2026",
    "Q4": "31/12/2026",
}

QUARTER_TAIL_LABELS_2026 = {
    "Q1": "16/03/2026..31/03/2026",
    "Q2": "16/06/2026..30/06/2026",
    "Q3": "16/09/2026..30/09/2026",
    "Q4": "16/12/2026..31/12/2026",
}

DATE_FORMAT = "%d/%m/%Y"


def parse_date(value: str) -> datetime.date:
    return datetime.strptime(value, DATE_FORMAT).date()


def parse_quarter_tail_label(label: str) -> Tuple[datetime.date, datetime.date]:
    start_str, end_str = label.split("..")
    return parse_date(start_str), parse_date(end_str)


def quarter_tail_date_range(quarter: str) -> Tuple[datetime.date, datetime.date]:
    return parse_quarter_tail_label(QUARTER_TAIL_LABELS_2026[quarter])


def sum_invoice_values_for_tail(
    invoices: Iterable[Mapping[str, Any]],
    quarter: str,
    date_key: str = "invoice_date",
    value_key: str = "invoice_value",
) -> float:
    start_date, end_date = quarter_tail_date_range(quarter)
    total = 0.0
    for invoice in invoices:
        invoice_date = parse_date(invoice[date_key])
        if start_date <= invoice_date <= end_date:
            total += float(invoice[value_key])
    return total


def next_period_label(selected_quarter: str) -> str:
    if selected_quarter == "Q4":
        return "2027"
    idx = QUARTER_ORDER.index(selected_quarter)
    return f"{QUARTER_ORDER[idx + 1]}-2026"


def build_customer_output_config(selected_quarter: str) -> dict:
    idx = QUARTER_ORDER.index(selected_quarter)
    active_quarters = QUARTER_ORDER[idx:]
    tail_labels = [QUARTER_TAIL_LABELS_2026[q] for q in active_quarters]
    next_label = next_period_label(selected_quarter)
    next_display_label = "2027" if selected_quarter == "Q4" else QUARTER_ORDER[idx + 1]

    return {
        "selected_quarter": selected_quarter,
        "active_quarters": active_quarters,
        "current_period_label": f"{selected_quarter}-2026",
        "current_pivot_label": f"{selected_quarter}-2026 - pivot",
        "percent_label": f"% for {selected_quarter}",
        "actual_label": f"Actual {selected_quarter}",
        "remaining_label": f"Remaining % from {selected_quarter.lower()}",
        "to_add_label": f"To add to {next_display_label}",
        "forecast_label": f"Forecast {next_display_label}",
        "next_period_label": next_label,
        "tail_labels": tail_labels,
        "later_quarter_labels": [f"{q}-2026" for q in active_quarters[1:]],
        "year_labels": ["2027", "2028", "2029", "2030"],
    }


def detect_selected_quarter_from_columns(columns) -> str:
    cols = set(columns)
    for quarter in QUARTER_ORDER:
        if f"% for {quarter}" in cols or f"Actual {quarter}" in cols:
            return quarter
    for quarter in QUARTER_ORDER:
        if f"{quarter}-2026 - pivot" in cols:
            return quarter
    return "Q1"