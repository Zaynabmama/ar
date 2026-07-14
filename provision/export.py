import io
import json
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import xlsxwriter

_TEMPLATE_PATH = Path(__file__).parent / "template_data.json"
_DATA = json.loads(_TEMPLATE_PATH.read_text(encoding="utf-8"))

ROW_TOKEN = _DATA["row_token"]    # substituted with the cell's Excel row
GATE_TOKEN = _DATA["gate_token"]  # substituted with $B{row-7} (master fill-down artifact, kept per user decision)
HEADER_ROW = _DATA["header_row"]              # Excel row 9
FIRST_DATA_ROW = _DATA["first_data_row"]      # Excel row 10
SUBTOTAL_ROW = _DATA["subtotal_row"]          # Excel row 7
SUBTOTAL_LAST = _DATA["subtotal_last_excel_row"]  # master hardcodes 19991
RATES_ROW = _DATA["rates_row"]                # Excel row 5

FIXED_TEXT_COLS = set(_DATA["text_cols"])
_STATIC_COLS = set(_DATA["static_cols"])


def _colnum(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _to_stored_form(formula: str) -> str:
    """Prefix Excel-365 functions with _xlfn. and LET/LAMBDA parameter names with
    _xlpm. as required by the xlsx storage format (Excel strips them for display)."""
    out = formula
    for fn in _DATA["future_funcs"]:
        out = re.sub(rf"(?<![A-Za-z0-9_.]){fn}\(", f"_xlfn.{fn}(", out)
    if "_xlfn.LET(" in out or "_xlfn.LAMBDA(" in out:
        for name in _DATA["param_names"]:
            out = re.sub(rf"(?<![A-Za-z0-9_.]){re.escape(name)}(?![A-Za-z0-9_.])", f"_xlpm.{name}", out)
    return out


# stored form is row-independent, so transform each template once at import
_STORED_TEMPLATES = {col: _to_stored_form(tpl) for col, tpl in _DATA["templates"].items()}
_STORED_LAMBDA = "=" + _to_stored_form(_DATA["lambda_formula"])


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
    """Build the AR Collection and Provision Forecast workbook (single ALL sheet,
    new Master File model: header row 9, data from row 10, MW_PROV_FC_CORE LAMBDA).

    df_fixed: one row per customer, columns keyed by output column letter
              (A-U static values incl. the K-O Not Due breakdown).
    """
    n_rows = len(df_fixed)
    last_data_row = FIRST_DATA_ROW + n_rows - 1

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"constant_memory": True})
    wb.define_name(_DATA["lambda_name"], _STORED_LAMBDA)
    fmts = _FormatCache(wb)
    row_styles = _DATA["styles"]["rows"]
    data_styles = _DATA["styles"]["data"]

    def _rowfmt(excel_row: int, col: str):
        st = row_styles.get(str(excel_row), {}).get(col)
        if not st:
            return None
        st = dict(st)
        nf = st.pop("num_format", None)
        return fmts.get(nf, st)

    ws = wb.add_worksheet("ALL")
    for idx, width in enumerate(_DATA["widths"]):
        ws.set_column(idx, idx, width)
    for r_str, height in _DATA["row_heights"].items():
        r = int(r_str)
        if r < FIRST_DATA_ROW:
            ws.set_row(r - 1, height)

    def _fill_styled_blanks(excel_row: int, written: set):
        for col, st in row_styles.get(str(excel_row), {}).items():
            if col not in written:
                stc = dict(st)
                nf = stc.pop("num_format", None)
                ws.write_blank(excel_row - 1, _colnum(col) - 1, None, fmts.get(nf, stc))

    # ---- title block (rows 1-3; B3 = AR Data Date, the model's control cell) ----
    for addr, text in _DATA["titles"].items():
        col, row = addr[0], int(addr[1:])
        if addr == "B3":
            continue  # master's own date; we write the picked ar_date below
        ws.write_string(row - 1, _colnum(col) - 1, text, _rowfmt(row, col))
    ws.write_datetime(2, 1, datetime(ar_date.year, ar_date.month, ar_date.day), _rowfmt(3, "B"))
    for r in (1, 2, 3):
        _fill_styled_blanks(r, {"A", "B"} if r == 3 else {"A"})

    # ---- provision rates (row 5) ----
    for col, rate in _DATA["rates"].items():
        ws.write_number(RATES_ROW - 1, _colnum(col) - 1, rate, _rowfmt(RATES_ROW, col))
    _fill_styled_blanks(RATES_ROW, set(_DATA["rates"]))

    # ---- month/label row (row 6, verbatim from master incl. its date serials) ----
    for col, cell in _DATA["row6_cells"].items():
        c = _colnum(col) - 1
        if cell["type"] == "number":
            ws.write_number(5, c, cell["value"], _rowfmt(6, col))
        else:
            ws.write_string(5, c, cell["value"], _rowfmt(6, col))
    _fill_styled_blanks(6, set(_DATA["row6_cells"]))

    # ---- totals (row 7, SUBTOTAL over the master's fixed 10:19991 range) ----
    for col in _DATA["subtotal_cols"]:
        ws.write_formula(SUBTOTAL_ROW - 1, _colnum(col) - 1,
                         f"=SUBTOTAL(9,{col}{FIRST_DATA_ROW}:{col}{SUBTOTAL_LAST})",
                         _rowfmt(SUBTOTAL_ROW, col))
    _fill_styled_blanks(SUBTOTAL_ROW, set(_DATA["subtotal_cols"]))

    # ---- headers (row 9, verbatim) ----
    for col, text in _DATA["headers"].items():
        ws.write_string(HEADER_ROW - 1, _colnum(col) - 1, text, _rowfmt(HEADER_ROW, col))
    _fill_styled_blanks(HEADER_ROW, set(_DATA["headers"]))

    # ---- data rows ----
    col_fmt = {}
    all_letters = (set(_DATA["num_formats"]) | set(data_styles) | _STATIC_COLS
                   | set(_DATA["input_cols_zero"]) | set(_DATA["input_cols_blank"])
                   | set(_DATA["manual_blank_cols"]))
    for col in all_letters:
        col_fmt[col] = fmts.get(num_format=_DATA["num_formats"].get(col), style=data_styles.get(col, {}))

    fixed_cols_present = [c for c in df_fixed.columns if c in _STATIC_COLS]
    gap_cols = [c for c in ("HB", "HC", "HD", "HE")
                if c in data_styles or c in _DATA["num_formats"]]
    blank_cols = _DATA["input_cols_blank"] + _DATA["manual_blank_cols"] + gap_cols

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

        for col in _DATA["input_cols_zero"]:
            ws.write_number(r, _colnum(col) - 1, 0, col_fmt.get(col))

        # blank, never 0: these cells drive IF(x="","",...) gates (Actual Collection)
        # or are manual entries (C BT, W/X prior provisions, AC Notes)
        for col in blank_cols:
            ws.write_blank(r, _colnum(col) - 1, None, col_fmt.get(col))

        gate_ref = f"$B{excel_row - 7}"
        for col, tpl in _STORED_TEMPLATES.items():
            f = tpl.replace(ROW_TOKEN, str(excel_row))
            if GATE_TOKEN in f:
                f = f.replace(GATE_TOKEN, gate_ref)
            ws.write_formula(r, _colnum(col) - 1, f, col_fmt.get(col))

    # ---- layout ----
    freeze_row, freeze_col = _DATA["freeze"]
    ws.freeze_panes(freeze_row, freeze_col)
    ws.autofilter(HEADER_ROW - 1, 0, max(last_data_row - 1, HEADER_ROW),
                  _colnum(_DATA["autofilter_last_col"]) - 1)
    for col in _DATA["hidden_cols"]:
        c = _colnum(col) - 1
        ws.set_column(c, c, None, None, {"hidden": True})

    wb.close()
    output.seek(0)
    return output.getvalue()
