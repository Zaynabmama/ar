"""Load the raw ORION/email exports.

Every ORION export has a variable-length metadata preamble (title, run date,
filter parameters) before the real data table, and some of them ship with a
broken <dimension> tag that makes naive openpyxl read-only iteration silently
truncate to column A. Reading through pandas.read_excel(engine="openpyxl")
sidesteps that, but we still don't know in advance how many preamble rows a
given export has, so every loader auto-detects the header row instead of
assuming a fixed offset.
"""

from __future__ import annotations

import pandas as pd

MAX_HEADER_SCAN_ROWS = 40


def detect_header_row(path, sheet_name, required_columns) -> int:
    """Scan the first rows of a sheet for the one that looks like the header.

    Returns a 0-based row index suitable for pandas' `header=` argument.
    """
    preview = pd.read_excel(
        path, sheet_name=sheet_name, header=None, nrows=MAX_HEADER_SCAN_ROWS, engine="openpyxl"
    )
    required = {c.strip().casefold() for c in required_columns}
    for row_idx in range(len(preview)):
        row_values = {
            str(v).strip().casefold() for v in preview.iloc[row_idx].tolist() if isinstance(v, str)
        }
        if required.issubset(row_values):
            return row_idx
    raise ValueError(
        f"Could not find a header row containing {sorted(required_columns)} "
        f"in the first {MAX_HEADER_SCAN_ROWS} rows of sheet {sheet_name!r}."
    )


def _load(path, sheet_name, required_columns) -> pd.DataFrame:
    header_row = detect_header_row(path, sheet_name, required_columns)
    # keep_default_na=False: pandas' default NA-string list includes "NA",
    # "N/A", "NULL", etc. The backlog report's Credit Status column uses the
    # literal text "NA" as a real status value (~13% of rows) -- with the
    # default behaviour pandas silently blanks it out before our filter ever
    # sees it. Genuinely empty cells still come through as NaN regardless of
    # this setting (there's no string to test against na_values for those).
    df = pd.read_excel(
        path, sheet_name=sheet_name, header=header_row, engine="openpyxl", keep_default_na=False, na_values=[]
    )
    df = df.replace("", pd.NA)
    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_ar_provision(path) -> pd.DataFrame:
    return _load(path, sheet_name=0, required_columns=["Cust Code", "Legal Entity", "Main Ac", "Ar Balance"])


def load_pdc(path) -> pd.DataFrame:
    return _load(path, sheet_name=0, required_columns=["Division", "Main Account", "Customer"])


def load_backlog(path) -> pd.DataFrame:
    return _load(path, sheet_name=0, required_columns=["Order", "Credit Status", "Customer Code"])


def load_insurance(path) -> pd.DataFrame:
    return _load(path, sheet_name=0, required_columns=["Customer Code", "Insurance Limit"])


def load_region_override_file(path) -> pd.DataFrame:
    """Loads a city/sub-region customer list (e.g. "UAE-AbuDhabi-AUH
    Customers list") -- same ORION export shape as the others. The
    'Addr State Code' column (e.g. 'AUH') is the override value: any
    customer listed here should map to that region instead of the country-
    level lookup. Returns columns ['Cust Code', 'Reporting Region']."""
    df = _load(path, sheet_name=0, required_columns=["Cust Code", "Addr State Code"])
    df = df.rename(columns={"Addr State Code": "Reporting Region"})
    return df[["Cust Code", "Reporting Region"]]
