import calendar

import pandas as pd

MONTH_ORDER_2026 = [f"{i:02d}" for i in range(1, 13)]


def _month_end_day(month: str, year: int = 2026) -> int:
    return calendar.monthrange(year, int(month))[1]


def month_end_label(month: str, year: int = 2026) -> str:
    """End-of-month label matching the provision tool's own month headers, e.g. '31-Jul-2026'."""
    day = _month_end_day(month, year)
    return f"{day:02d}-{calendar.month_abbr[int(month)]}-{year}"


MONTH_LABELS_2026 = {m: month_end_label(m) for m in MONTH_ORDER_2026}

MONTH_SELECT_LABELS = {
    m: f"{calendar.month_name[int(m)]} ({MONTH_LABELS_2026[m]})" for m in MONTH_ORDER_2026
}


def month_date_range(month: str, year: int = 2026) -> tuple[pd.Timestamp, pd.Timestamp]:
    start = pd.Timestamp(year, int(month), 1)
    end = pd.Timestamp(year, int(month), _month_end_day(month, year))
    return start, end


def build_customer_output_config(selected_month: str) -> dict:
    idx = MONTH_ORDER_2026.index(selected_month)
    active_months = MONTH_ORDER_2026[idx:]
    return {
        "selected_month": selected_month,
        "active_months": active_months,
        "month_labels": [MONTH_LABELS_2026[m] for m in active_months],
        "year_labels": ["2027", "2028", "2029", "2030"],
    }
