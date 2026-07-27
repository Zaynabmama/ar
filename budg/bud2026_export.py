"""BUD2026 quarterly model export.

Rebuilds the "AR Collection and Provision Forecast - Master quarterly" ALL
sheet from bud_rows (see bud2026_mapper). Layout, formulas and styling come
from bud2026_template.json, extracted verbatim from the master workbook:

  rows 1-4  titles - "Period ended" is derived from the AR Data Date
  row 5     AR Data Date control cell (B5) - gates the quarter formulas
  row 7     live provision rates J7:U7 (12 columns - Not Due breakdown feeds
            the quarters-elapsed rate lookup used by AR Provision FC/Actual)
  row 8     "Balance at ..." band + Q1-Q4 banner band
  row 9     SUBTOTAL row over the full data range (skips the Y spacer and
            Notes columns, which the master leaves untotaled)
  row 11    headers, data from row 12
  per row   28 live formula columns (Z:AC base provisions + 6 per quarter);
            8 of those (AR Provision FC/Actual AR Provision x4 quarters) are
            single-cell array formulas in the master, written accordingly
  column Y  blank spacer column between the manual W/X provision snapshots
            and the Z:AC base-provision block - purely visual, never written
  inputs    Collections FC (pre-filled when mapped), Specific Alloc, Actual
            Collection - written blank, guarded by the master's two
            "no double counting" data validations (one per quarter block,
            formulas re-anchored per range)
"""
import datetime as _dt
import io
import json
import os
import re

import numpy as np
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

from budg.bud2026_headers import COLLECTION_FC_COLUMNS, QUARTER_ENDS_2026, VALUE_COLUMNS

_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "bud2026_template.json")
_template_cache = None

HEADER_ROW = 11          # 1-based Excel rows
DATA_START_ROW = 12
SUBTOTAL_ROW = 9

# quarter block anchor columns (first column of each 15-column block)
_QUARTER_ANCHORS = ["AE", "AT", "BI", "BX"]
_BLOCK_WIDTH = 15


def _load_template() -> dict:
    global _template_cache
    if _template_cache is None:
        with open(_TEMPLATE_PATH, encoding="utf-8") as f:
            _template_cache = json.load(f)
    return _template_cache


def _col_idx(letter: str) -> int:
    """'A' -> 0"""
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


_HALIGN = {-4108: "center", -4131: "left", -4152: "right", 7: "center_across"}


def _fmt_props(style: dict | None, num_format: str | None = None, *, bold=None, wrap=None) -> dict:
    props = {}
    if num_format and num_format != "General":
        props["num_format"] = num_format
    if style:
        if style.get("fill"):
            props["bg_color"] = style["fill"]
        if style.get("font_color") and style["font_color"] != "#000000":
            props["font_color"] = style["font_color"]
        if style.get("bold"):
            props["bold"] = True
        if style.get("wrap"):
            props["text_wrap"] = True
        halign = _HALIGN.get(style.get("halign"))
        if halign:
            props["align"] = halign
        if style.get("border_bottom"):
            props["bottom"] = 1
    if bold is not None:
        props["bold"] = bold
    if wrap is not None:
        props["text_wrap"] = wrap
    return props


def _safe_value(value):
    if value is None:
        return None
    if isinstance(value, str):
        return value
    if pd.isna(value):
        return None
    if isinstance(value, (float, np.floating)) and not np.isfinite(value):
        return None
    return value


def _period_ended_title(ar_date: _dt.date) -> str:
    """Latest FY2026 quarter end strictly before the AR Data Date (the master
    shows 'Period ended March 31, 2026' with B5 = 30-Jun-2026)."""
    ends = [_dt.date(2026, m, d) for m, d in QUARTER_ENDS_2026]
    before = [e for e in ends if e < ar_date]
    period = before[-1] if before else _dt.date(2025, 12, 31)
    return "Period ended " + period.strftime("%B %d, %Y").replace(" 0", " ")


def export_bud2026_quarterly(bud_rows: pd.DataFrame, ar_date: _dt.date) -> bytes:
    tpl = _load_template()
    styles = tpl["styles"]
    ncols = tpl["ncols"]
    headers = tpl["headers"]
    formulas = tpl["formulas"]
    numfmt12 = tpl["numfmt_row12"]

    n_rows = len(bud_rows)
    last_row = DATA_START_ROW + max(n_rows, 1) - 1

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"constant_memory": True})
    ws = wb.add_worksheet("ALL")

    fmt_cache: dict[tuple, object] = {}

    def fmt(props: dict):
        key = tuple(sorted(props.items()))
        if key not in fmt_cache:
            fmt_cache[key] = wb.add_format(props)
        return fmt_cache[key]

    letters = [xl_col_to_name(c) for c in range(ncols)]

    # ---- per-column formats ----
    header_fmts, data_fmts, subtotal_fmts = [], [], []
    for c, letter in enumerate(letters):
        hs = dict(styles["header_styles"][c])
        hs["wrap"] = True
        header_fmts.append(fmt(_fmt_props(hs) | {"valign": "vcenter"}))
        ds = styles["data_styles"][c]
        props = _fmt_props(None, numfmt12.get(letter))
        if ds.get("fill"):
            props["bg_color"] = ds["fill"]
        data_fmts.append(fmt(props) if props else None)
        ss = dict(styles["subtotal_styles"][c])
        ss["halign"] = None
        sub_num = tpl["numfmt_cells"]["AE9" if c >= _col_idx("AE") else "I9"]
        subtotal_fmts.append(fmt(_fmt_props(ss, sub_num)))

    # ---- columns: widths + outline groups ----
    grouped = {}
    for g in tpl["column_groups"]:
        for c in range(g["min"], g["max"] + 1):
            grouped[c - 1] = {"level": g["level"], "hidden": g["hidden"]}
    collapsed_cols = {g["max"] for g in tpl["column_groups"] if g["hidden"]}
    for c in range(ncols):
        options = dict(grouped.get(c, {}))
        if c - 1 in {m for m in collapsed_cols}:
            options["collapsed"] = True
        ws.set_column(c, c, styles["col_widths"][c], None, options or None)

    # ---- rows 1-5: titles + AR Data Date ----
    title_fmt = fmt({"bold": True})
    titles = tpl["titles"]
    ws.write_string(0, 0, titles["1"], title_fmt)
    ws.write_string(1, 0, titles["2"], title_fmt)
    ws.write_string(2, 0, _period_ended_title(ar_date), title_fmt)
    ws.write_string(3, 0, titles["4"], title_fmt)
    ws.write_string(4, 0, tpl["ar_date_label"], title_fmt)
    ws.write_datetime(
        4, 1, _dt.datetime.combine(ar_date, _dt.time()),
        fmt({"num_format": tpl["numfmt_cells"]["B5"], "align": "left"}),
    )

    # ---- row 7: provision rates ----
    rate_fmt = fmt({"num_format": tpl["numfmt_cells"]["J7"], "bottom": 1})
    for c in range(_col_idx("J"), _col_idx("U") + 1):
        rate = tpl["rates"].get(letters[c])
        if rate is not None:
            ws.write_number(6, c, rate, rate_fmt)
        else:
            ws.write_blank(6, c, None, rate_fmt)

    # ---- row 8: balance label band + quarter banner band ----
    row8_labels = tpl["row8_labels"]
    dark_band = fmt(_fmt_props(styles["cells"]["I8"]))
    for c in range(_col_idx("I"), _col_idx("T") + 1):
        if c == _col_idx("I"):
            ws.write_string(7, c, row8_labels["I"], dark_band)
        else:
            ws.write_blank(7, c, None, dark_band)
    blue_band = fmt(_fmt_props(styles["cells"]["AE8"]) | {"align": "center_across"})
    for c in range(_col_idx("AE"), ncols):
        label = row8_labels.get(letters[c])
        if label:
            ws.write_string(7, c, label, blue_band)
        else:
            ws.write_blank(7, c, None, blue_band)

    # ---- row 9: subtotals over the full data range ----
    for c in range(_col_idx("I"), ncols):
        letter = letters[c]
        if letter in ("Y", "AD"):   # spacer column / Notes column - no subtotal in the master
            continue
        ws.write_formula(
            SUBTOTAL_ROW - 1, c,
            f"=SUBTOTAL(9,{letter}{DATA_START_ROW}:{letter}{last_row})",
            subtotal_fmts[c],
        )

    # ---- row 11: headers ----
    for c, header in enumerate(headers):
        ws.write(HEADER_ROW - 1, c, header, header_fmts[c])
    ws.set_row(HEADER_ROW - 1, styles["row_heights"]["11"])

    # ---- data rows ----
    value_cols = {_col_idx(letter): name for name, letter in VALUE_COLUMNS.items()}
    collection_cols = {_col_idx(letter): name for name, letter in COLLECTION_FC_COLUMNS.items()}
    formula_cols = {_col_idx(letter): tmpl for letter, tmpl in formulas.items()}
    # the master enters these as single-cell (legacy CSE) array formulas
    array_formula_cols = {_col_idx(letter) for letter in tpl.get("array_formula_cols", [])}
    main_ac_idx = _col_idx(VALUE_COLUMNS["Main Ac"])
    digits = re.compile(r"^-?\d+$")

    records = bud_rows.to_dict("records")
    for i, record in enumerate(records):
        r = DATA_START_ROW - 1 + i          # 0-based sheet row
        excel_row = str(r + 1)
        for c in range(ncols):
            cfmt = data_fmts[c]
            if c in formula_cols:
                formula = formula_cols[c].replace("«R»", excel_row)
                if c in array_formula_cols:
                    ws.write_dynamic_array_formula(r, c, r, c, formula, cfmt)
                else:
                    ws.write_formula(r, c, formula, cfmt)
                continue
            if c in value_cols:
                value = _safe_value(record.get(value_cols[c]))
                # Main Ac must be numeric: the formulas compare $G12<>12301
                if c == main_ac_idx and isinstance(value, str) and digits.match(value.strip()):
                    value = int(value.strip())
                if value is None or value == "":
                    ws.write_blank(r, c, None, cfmt)
                elif isinstance(value, str):
                    ws.write_string(r, c, value, cfmt)
                else:
                    ws.write_number(r, c, float(value), cfmt)
                continue
            if c in collection_cols:
                value = _safe_value(record.get(collection_cols[c]))
                if value is not None and not isinstance(value, str) and float(value) != 0.0:
                    ws.write_number(r, c, float(value), cfmt)
                else:
                    ws.write_blank(r, c, None, cfmt)
                continue
            # manual inputs (W, X, AC, Specific Alloc, Actual Collection, ...)
            if cfmt is not None:
                ws.write_blank(r, c, None, cfmt)

    # ---- data validation: "no double counting", re-anchored per block ----
    for anchor in _QUARTER_ANCHORS:
        a = _col_idx(anchor)
        spec_first, spec_last = letters[a + 1], letters[a + 7]
        ws.data_validation(
            f"{anchor}{DATA_START_ROW}:{anchor}{last_row}",
            {
                "validate": "custom",
                "value": f"=SUM({spec_first}{DATA_START_ROW}:{spec_last}{DATA_START_ROW})=0",
                "error_title": "Avoid double-counting",
                "error_message": "Specific allocations already exist for this row/quarter.",
            },
        )
        ws.data_validation(
            f"{spec_first}{DATA_START_ROW}:{spec_last}{last_row}",
            {
                "validate": "custom",
                "value": f"={anchor}{DATA_START_ROW}=0",
                "error_title": "Avoid double-counting",
                "error_message": "FIFO collection already exists for this row/quarter.",
            },
        )

    ws.freeze_panes(DATA_START_ROW - 1, 2)   # C12
    ws.autofilter(HEADER_ROW - 1, 0, last_row - 1, ncols - 1)

    wb.close()
    output.seek(0)
    return output.getvalue()
