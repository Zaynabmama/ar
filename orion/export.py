import io

import pandas as pd
import xlsxwriter

from common.quarter_utils import build_customer_output_config

ZERO_COLLECTION_MAIN_ACCOUNTS = ("12302", "12304", "12306")


def num_to_col_letters(n: int) -> str:
    s = ""
    n += 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def build_col_map(df) -> dict:
    return {col: num_to_col_letters(i) for i, col in enumerate(df.columns)}


def normalize_all_date_strings(df: pd.DataFrame) -> pd.DataFrame:
    import re

    if df is None or df.empty:
        return df
    df = df.copy()
    iso = re.compile(r"^\s*\d{4}-\d{2}-\d{2}")

    def _looks_like_iso_date(value) -> bool:
        if pd.isna(value):
            return False
        text = str(value).strip()
        if not text or text.lower() == "nan":
            return False
        return bool(iso.match(text))

    for col in df.columns:
        if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
            s = df[col].astype("string")
            if not s.head(50).map(_looks_like_iso_date).any():
                continue
            s = (
                s.replace("\u00A0", " ", regex=False)
                .str.replace(r"[^\x00-\x7F]", "", regex=True)
                .str.strip()
            )
            s = s.where(~s.str.match(r"^\d{4}-\d{2}-\d{2}"), s.str.slice(0, 10))
            df[col] = s
    return df


def coerce_export_dates(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    df = df.copy()
    date_like_columns = [
        "Document Date",
        "Document Due Date",
    ]
    for col in date_like_columns:
        if col not in df.columns:
            continue
        raw = df[col]
        cleaned = raw.astype(str).str.strip().replace({"nan": "", "NaT": "", "": pd.NA})
        dt = pd.to_datetime(cleaned, errors="coerce")
        df[col] = dt.where(dt.notna(), raw)
    return df


def fast_excel_download_multiple_with_formulas(
    df_main, df_customer, df_invoice, selected_quarter="Q1"
) -> io.BytesIO:
    cfg = build_customer_output_config(selected_quarter)

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"constant_memory": True})
    header_fmt = wb.add_format({"bold": True, "bg_color": "#F2F2F2"})
    date_fmt = wb.add_format({"num_format": "dd/mm/yyyy"})

    def write_sheet(ws, df):
        prepared = coerce_export_dates(normalize_all_date_strings(df.copy())).fillna("")
        for c_idx, name in enumerate(prepared.columns):
            ws.write(0, c_idx, str(name), header_fmt)
        for r_idx, row in enumerate(prepared.itertuples(index=False), start=1):
            for c_idx, value in enumerate(row):
                header = prepared.columns[c_idx]
                if header in {"Document Date", "Document Due Date"}:
                    if pd.notna(value) and hasattr(value, "to_pydatetime"):
                        ws.write_datetime(r_idx, c_idx, value.to_pydatetime(), date_fmt)
                    else:
                        ws.write(r_idx, c_idx, "")
                else:
                    ws.write(r_idx, c_idx, value)
        return prepared

    ws_main = wb.add_worksheet("AR_Backlog")
    main_df = write_sheet(ws_main, df_main)

    ws_cust = wb.add_worksheet("By_Customer")
    cust_df = normalize_all_date_strings(df_customer.copy()).fillna("")
    for c_idx, name in enumerate(cust_df.columns):
        ws_cust.write(0, c_idx, str(name), header_fmt)

    col_map = build_col_map(cust_df)

    def idx(name):
        return list(cust_df.columns).index(name)

    for r_idx, row in enumerate(cust_df.to_numpy(), start=1):
        excel_row = r_idx + 1
        ws_cust.write_row(r_idx, 0, row.tolist())

        current_period = col_map.get(cfg["current_pivot_label"])
        total = col_map.get(cfg["total_label"])
        forecasted = col_map.get(cfg["forecasted_label"])
        remaining = col_map.get(cfg["remaining_label"])
        next_period = col_map.get(cfg["next_period_label"])
        next_period_tail = col_map.get(cfg["next_period_tail_label"]) if cfg["next_period_tail_label"] else None
        main_ac = col_map.get("Main Ac")
        current_period_output = col_map.get(cfg["current_period_label"])
        overdue = col_map.get("Overdue")
        on_account = col_map.get("On account")
        ageing_over_365 = col_map.get("Ageing > 365")

        collection_zero_guard = None
        if main_ac:
            guards = [f'${main_ac}{excel_row}="{value}"' for value in ZERO_COLLECTION_MAIN_ACCOUNTS]
            collection_zero_guard = f"OR({','.join(guards)})"

        if current_period_output and current_period and overdue and on_account and ageing_over_365:
            period_formula = (
                f"IF((${current_period}{excel_row}+${overdue}{excel_row}+${on_account}{excel_row}"
                f"-${ageing_over_365}{excel_row})>0,"
                f"${current_period}{excel_row}+${overdue}{excel_row}+${on_account}{excel_row}"
                f"-${ageing_over_365}{excel_row},0)"
            )
            if collection_zero_guard:
                period_formula = f"IF({collection_zero_guard},0,{period_formula})"
            ws_cust.write_formula(
                r_idx,
                idx(cfg["current_period_label"]),
                f"={period_formula}",
            )

        # Total Q[n] and Forecasted Q[n] collection are already written as
        # plain values by write_row above (see orion/processor.py) - Total is
        # tool-computed, Forecasted is seeded from the tail amount and meant
        # to be overwritten by hand. Only Remaining and Forecast [next]
        # Collection stay as live formulas so they track manual edits.
        if total and forecasted:
            remaining_formula = f"IFERROR(${total}{excel_row}-${forecasted}{excel_row},0)"
            if collection_zero_guard:
                remaining_formula = f"IF({collection_zero_guard},0,{remaining_formula})"
            ws_cust.write_formula(
                r_idx,
                idx(cfg["remaining_label"]),
                f"={remaining_formula}",
            )

        if remaining and next_period:
            forecast_parts = f"${remaining}{excel_row}+${next_period}{excel_row}"
            if next_period_tail:
                forecast_parts += f"+${next_period_tail}{excel_row}"
            forecast_formula = f"IFERROR({forecast_parts},0)"
            if collection_zero_guard:
                forecast_formula = f"IF({collection_zero_guard},0,{forecast_formula})"
            ws_cust.write_formula(
                r_idx,
                idx(cfg["forecast_label"]),
                f"={forecast_formula}",
            )

    ws_cust.freeze_panes(1, 0)
    ws_cust.autofilter(0, 0, max(1, len(cust_df)), len(cust_df.columns) - 1)

    ws_inv = wb.add_worksheet("Invoice")
    inv_df = write_sheet(ws_inv, df_invoice)

    wb.close()
    output.seek(0)
    return output


def fast_excel_download_monthly(df_main, df_customer, df_invoice, selected_month="01") -> io.BytesIO:
    """Monthly counterpart of fast_excel_download_multiple_with_formulas.

    By_Customer is buckets-only (no %/Actual/Remaining/Forecast mechanic), so
    unlike the quarterly export it needs no per-row write_formula calls - it's
    written the same plain way as AR_Backlog/Invoice.
    """
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"constant_memory": True})
    header_fmt = wb.add_format({"bold": True, "bg_color": "#F2F2F2"})
    date_fmt = wb.add_format({"num_format": "dd/mm/yyyy"})

    def write_sheet(ws, df):
        prepared = coerce_export_dates(normalize_all_date_strings(df.copy())).fillna("")
        for c_idx, name in enumerate(prepared.columns):
            ws.write(0, c_idx, str(name), header_fmt)
        for r_idx, row in enumerate(prepared.itertuples(index=False), start=1):
            for c_idx, value in enumerate(row):
                header = prepared.columns[c_idx]
                if header in {"Document Date", "Document Due Date"}:
                    if pd.notna(value) and hasattr(value, "to_pydatetime"):
                        ws.write_datetime(r_idx, c_idx, value.to_pydatetime(), date_fmt)
                    else:
                        ws.write(r_idx, c_idx, "")
                else:
                    ws.write(r_idx, c_idx, value)
        return prepared

    ws_main = wb.add_worksheet("AR_Backlog")
    write_sheet(ws_main, df_main)

    ws_cust = wb.add_worksheet("By_Customer")
    cust_df = write_sheet(ws_cust, df_customer)
    ws_cust.freeze_panes(1, 0)
    ws_cust.autofilter(0, 0, max(1, len(cust_df)), len(cust_df.columns) - 1)

    ws_inv = wb.add_worksheet("Invoice")
    write_sheet(ws_inv, df_invoice)

    wb.close()
    output.seek(0)
    return output
