import io
import json
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import xlsxwriter

_TEMPLATE_PATH = Path(__file__).parent / "template_data.json"
_DATA = json.loads(_TEMPLATE_PATH.read_text(encoding="utf-8"))

ROW_TOKEN = _DATA["row_token"]
HEADER_ROW = _DATA["header_row"]          # Excel row 11
FIRST_DATA_ROW = _DATA["first_data_row"]  # Excel row 12

MONTH_ENDS_2026 = [
    date(2026, 1, 31), date(2026, 2, 28), date(2026, 3, 31), date(2026, 4, 30),
    date(2026, 5, 31), date(2026, 6, 30), date(2026, 7, 31), date(2026, 8, 31),
    date(2026, 9, 30), date(2026, 10, 31), date(2026, 11, 30), date(2026, 12, 31),
]

# fixed-value columns fed by the mapper (letters); L-P only in fallback mode
FIXED_TEXT_COLS = {"A", "B", "C", "D", "E", "F", "H"}


def _colnum(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _to_stored_form(formula: str) -> str:
    """Prefix Excel-365 functions with _xlfn. and LET variable names with _xlpm.
    as required by the xlsx storage format (Excel strips them for display)."""
    out = formula
    for fn in _DATA["future_funcs"]:
        out = re.sub(rf"(?<![A-Za-z0-9_.]){fn}\(", f"_xlfn.{fn}(", out)
    if "_xlfn.LET(" in out:
        for name in _DATA["let_var_names"]:
            out = re.sub(rf"(?<![A-Za-z0-9_.]){re.escape(name)}(?![A-Za-z0-9_.])", f"_xlpm.{name}", out)
    return out


# stored form is row-independent, so transform each template once at import
_STORED_TEMPLATES = {col: _to_stored_form(tpl) for col, tpl in _DATA["templates"].items()}


class _FormatCache:
    def __init__(self, wb):
        self.wb = wb
        self.cache = {}

    def get(self, num_format=None, style=None):
        style = style or {}
        key = (num_format, tuple(sorted(style.items())))
        if key not in self.cache:
            props = {}
            if num_format and num_format != "General":
                props["num_format"] = num_format
            if style.get("fill"):
                props["bg_color"] = style["fill"]
            if style.get("bold"):
                props["bold"] = True
            if style.get("fcolor"):
                props["font_color"] = style["fcolor"]
            if style.get("wrap"):
                props["text_wrap"] = True
            if style.get("halign"):
                props["align"] = style["halign"]
            if style.get("valign"):
                props["valign"] = style["valign"]
            self.cache[key] = self.wb.add_format(props)
        return self.cache[key]


def export_provision_forecast(df_fixed: pd.DataFrame, ar_date: date) -> bytes:
    """Build the AR Collection and Provision Forecast workbook (single ALL sheet).

    df_fixed: one row per customer, columns keyed by output column letter (A-Y,
              including the L-P Not Due breakdown values).
    """
    n_rows = len(df_fixed)
    last_data_row = FIRST_DATA_ROW + n_rows - 1
    n_cols = len(_DATA["widths"])

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"constant_memory": True})
    fmts = _FormatCache(wb)
    styles = _DATA.get("styles", {})

    ws = wb.add_worksheet("ALL")
    for idx, width in enumerate(_DATA["widths"]):
        ws.set_column(idx, idx, width)

    # ---- title block (rows 1-5) ----
    title_fmt = fmts.get(style=styles.get("title", {"bold": True}))
    ws.write(0, 0, "Mindware", title_fmt)
    ws.write(1, 0, "QBR Budget - FY 2026", title_fmt)
    ws.write(2, 0, f"Period ended {ar_date:%B %d, %Y}", title_fmt)
    ws.write(3, 0, "USD", title_fmt)
    ws.write(4, 0, "AR Data Date:", fmts.get(style=styles.get("row5_a", {"bold": True})))
    ws.write_datetime(
        4, 1, datetime(ar_date.year, ar_date.month, ar_date.day),
        fmts.get(num_format="dd-mmm-yyyy", style=styles.get("row5_b", {})),
    )

    # ---- rates (row 7) ----
    rate_fmt = fmts.get(num_format="0%", style=styles.get("row7", {}))
    for col, rate in _DATA["rates"].items():
        ws.write_number(6, _colnum(col) - 1, rate, rate_fmt)

    # ---- month header dates (row 8, centered across each monthly block) ----
    ws.write(7, _colnum("J") - 1, f"Balance at {ar_date:%d-%m-%Y}",
             fmts.get(style=styles.get("row8", {}).get("J", {})))
    for span, month_end in zip(_DATA["month_spans"], MONTH_ENDS_2026):
        start, end = re.findall(r"([A-Z]+)8", span)
        month_style = dict(styles.get("row8", {}).get(start, {}))
        month_style["halign"] = "center_across"
        month_fmt = fmts.get(num_format="mmm-yy", style=month_style)
        ws.write_datetime(7, _colnum(start) - 1,
                          datetime(month_end.year, month_end.month, month_end.day), month_fmt)
        for c in range(_colnum(start), _colnum(end)):
            ws.write_blank(7, c, None, month_fmt)

    # ---- totals (row 9, SUBTOTAL so filtering works) ----
    for col in _DATA["row9_cols"]:
        fmt = fmts.get(num_format=_DATA["row9_num_formats"].get(col),
                       style=styles.get("row9", {}).get(col, {}))
        ws.write_formula(8, _colnum(col) - 1, f"=SUBTOTAL(9,{col}{FIRST_DATA_ROW}:{col}{last_data_row})", fmt)

    # ---- headers (row 11) ----
    headers = dict(_DATA["headers"])
    headers["X"] = f"AR Provision at\n{ar_date:%d-%m-%Y}"
    ws.set_row(HEADER_ROW - 1, _DATA.get("row11_height", 57))
    for col, text in headers.items():
        fmt = fmts.get(style=styles.get("row11", {}).get(col, {"bold": True, "wrap": True}))
        ws.write(HEADER_ROW - 1, _colnum(col) - 1, text, fmt)

    # ---- data rows ----
    data_styles = styles.get("data", {})
    col_fmt = {}
    all_letters = set(_DATA["num_formats"]) | set(data_styles) | set(_DATA["input_cols"]) | set("ABCDEFHIJKLMNOPQRSTUVWXY")
    for col in all_letters:
        col_fmt[col] = fmts.get(num_format=_DATA["num_formats"].get(col), style=data_styles.get(col, {}))

    fixed_cols_present = [c for c in df_fixed.columns if isinstance(c, str) and c.isalpha()]

    for i in range(n_rows):
        excel_row = FIRST_DATA_ROW + i
        r = excel_row - 1
        row = df_fixed.iloc[i]

        for col in fixed_cols_present:
            value = row[col]
            c = _colnum(col) - 1
            fmt = col_fmt.get(col)
            if value is None or (isinstance(value, float) and pd.isna(value)) or value == "":
                ws.write_blank(r, c, None, fmt)
            elif col in FIXED_TEXT_COLS:
                ws.write_string(r, c, str(value), fmt)
            elif col == "G":
                # Main Ac must be numeric where possible: formulas test G=12301
                try:
                    ws.write_number(r, c, float(str(value)), fmt)
                except ValueError:
                    ws.write_string(r, c, str(value), fmt)
            else:
                try:
                    ws.write_number(r, c, float(value), fmt)
                except (TypeError, ValueError):
                    ws.write_string(r, c, str(value), fmt)

        for col in _DATA["input_cols"]:
            ws.write_number(r, _colnum(col) - 1, 0, col_fmt.get(col))

        for col, tpl in _STORED_TEMPLATES.items():
            ws.write_formula(r, _colnum(col) - 1, tpl.replace(ROW_TOKEN, str(excel_row)), col_fmt.get(col))

    # ---- layout ----
    ws.freeze_panes(HEADER_ROW, 2)
    ws.autofilter(HEADER_ROW - 1, 0, max(last_data_row - 1, HEADER_ROW), n_cols - 1)

    wb.close()
    output.seek(0)
    return output.getvalue()
