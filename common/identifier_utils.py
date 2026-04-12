import math
import re
from numbers import Integral, Real

import pandas as pd


_INTEGER_DECIMAL_PATTERN = re.compile(r"^[+-]?\d+\.0+$")
_EMPTY_TOKENS = {"", "nan", "nat", "none", "<na>"}


def normalize_excel_identifier(value) -> str:
    """
    Normalize Excel-style identifier cells without changing real text codes.

    Examples:
    - 12301.0 -> "12301"
    - "12301.00" -> "12301"
    - "00123" -> "00123"
    - NaN -> ""
    """
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass

    if isinstance(value, str):
        text = value.replace("\u00A0", " ").strip()
        if text.lower() in _EMPTY_TOKENS:
            return ""
        if _INTEGER_DECIMAL_PATTERN.fullmatch(text):
            return text.split(".", 1)[0]
        return text

    if isinstance(value, bool):
        return str(value).strip()

    if isinstance(value, Integral):
        return str(int(value))

    if isinstance(value, Real):
        number = float(value)
        if not math.isfinite(number):
            return ""
        if number.is_integer():
            return str(int(number))
        return format(number, "g")

    text = str(value).replace("\u00A0", " ").strip()
    return "" if text.lower() in _EMPTY_TOKENS else text


def normalize_excel_identifier_series(series: pd.Series) -> pd.Series:
    return series.apply(normalize_excel_identifier)
