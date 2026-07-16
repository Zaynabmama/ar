"""Per-BUM output files (stage: split).

Computes the enriched master in Python (same math as the Sheet1 formulas,
verified cell-for-cell against the user's manual working file), then builds
one Excel file per BUM per the filter rules:

  * "CUS" sheet  - by customer (or customer+brand), sums of buckets /
    On Account / Not Due / Additional due, max of Insurance, Backlog,
    PDC, LC&BG; Ar Balance / Overdue as live formulas; SUBTOTAL row and
    ratio row on top. Sorted by Overdue descending.
  * "Inv" sheet  - the filtered invoice rows; AR Balance as formula;
    SUBTOTAL row; S1 = Inv total - CUS total (cross-check, must be 0).
  * Ehab's file additionally gets both CUS variants and the insurance
    report (limit - backlog - Net AR = available limit per customer).

Backlog rows whose Credit Status is 'NA' are dropped before pivoting.
"""

from __future__ import annotations

import calendar
import io
import zipfile
from datetime import date, datetime

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

from BUM.logic import _pivot, _read_csv, _read_main, _Report

_FMT_DATE = "mm-dd-yy"
_FMT_ACC = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)'

CUS_LOOKUP_COLS = [("Insurance", "insurance"), ("Backlog", "backlog"),
                   ("PDC2", "pdc"), ("LC/BG", "lcbg")]
EHAB_LOOKUP_COLS = [("Max Insurance", "insurance"), ("Backlog", "backlog"),
                    ("LC/BG", "lcbg"), ("PDC", "pdc")]

_QG = ("QNAL", "GCC")

# (file label, filter(row) -> bool, by-customer includes brands)
FILE_SPECS = [
    ("Anjali", lambda r: r["region"] in _QG and r["bum"] == "Anjali", True),
    ("Prashant", lambda r: r["region"] in _QG and r["bum"] == "Prashant", True),
    ("Sofia", lambda r: r["region"] in _QG and r["bum"] == "SOFIA", True),
    ("Fahad", lambda r: r["region"] in _QG and r["bum"] == "Fahad", True),
    ("Hassan", lambda r: r["region"] in _QG and r["bum"] == "Hassan", True),
    ("Neelu", lambda r: r["region"] in _QG and r["bum"] == "Neelu", True),
    ("Renewals", lambda r: r["region"] in _QG and r["renewals"] == "Renewals", True),
    ("Autodesk", lambda r: r["bum"] == "Brent", False),
    ("Dell", lambda r: r["region"] in _QG and r["bum"] == "Tarek", False),
    ("Lenovo", lambda r: r["region"] in _QG and r["bum"] == "Jinesh & Prashant", False),
    ("GSI", lambda r: r["shelly"] == "GSI", False),
    ("AUH", lambda r: r["region"] in _QG and r["auh"] == "AUH", False),
    ("Bahrain Oman Yemen", lambda r: r["cust_region"] in ("BAHRAIN", "OMAN", "YEMEN"), False),
    ("Kuwait", lambda r: r["cust_region"] == "KUWAIT", False),
    ("Pakistan Afghanistan", lambda r: r["cust_region"] in ("PAKISTAN", "AFGHANISTAN"), False),
    ("SE Africa", lambda r: r["seafrica"] == "SE Africa", False),
]


def _up(v) -> str:
    return str(v).strip().upper() if v not in (None, "") else ""


def compute_dataset(
    main_bytes: bytes,
    auh_bytes: bytes,
    renewals_bytes: bytes,
    insurance_bytes: bytes,
    pdc_bytes: bytes,
    backlog_bytes: bytes,
):
    """Parse the six inputs and compute every derived column per invoice row."""
    as_of, headers, data = _read_main(main_bytes)
    idx = {h: i for i, h in enumerate(headers)}

    auh = _Report(auh_bytes, "Cust Code", "AUH")
    auh_set = {
        _up(row[0])
        for row in auh.raw[auh.first_data_row - 1 : auh.last_data_row]
        if row and row[0] not in (None, "")
    }
    renewals = _Report(renewals_bytes, "Invoice number", "Renewals")
    renew_set = {
        _up(row[0])
        for row in renewals.raw[renewals.first_data_row - 1 : renewals.last_data_row]
        if row and row[0] not in (None, "")
    }
    insurance = _Report(insurance_bytes, "Customer Code", "Insurance")
    ih = {h: i for i, h in enumerate(insurance.headers)}
    ins_map: dict[str, float] = {}
    ins_rows = []  # (code, name, country, limit) first occurrence per code
    for row in insurance.raw[insurance.first_data_row - 1 : insurance.last_data_row]:
        if not row or row[0] in (None, ""):
            continue
        code = _up(row[0])
        if code in ins_map:
            continue
        limit = float(row[ih["Insurance Limit"]] or 0)
        ins_map[code] = limit
        ins_rows.append(
            (row[0], row[ih["Customer name"]], row[ih["Region Name"]], limit)
        )

    pdc_rows, _, _ = _pivot(_Report(pdc_bytes, "Division", "PDC"), "Sub Account", "LC Amount")
    pdc_map = {_up(k): v for k, v in pdc_rows}

    backlog = _Report(backlog_bytes, "Order", "Backlog")
    cs = backlog.headers.index("Credit Status")
    backlog.raw = [
        row
        for i, row in enumerate(backlog.raw)
        if i < backlog.first_data_row - 1
        or not row
        or len(row) <= cs
        or _up(row[cs]) != "NA"
    ]
    backlog.last_data_row = len(backlog.raw)
    bkl_rows, _, _ = _pivot(backlog, "Customer Code", "Pending Val (Lc)")
    bkl_map = {_up(k): v for k, v in bkl_rows}

    bum_map = {_up(r[0]): r[1] for r in _read_csv("bum_fixed.csv")[1:]}
    reg_map = {_up(r[0]): r[1] for r in _read_csv("region.csv")[1:]}
    sea_map = {_up(r[0]): r[1] for r in _read_csv("se_africa.csv")[1:]}
    gsi_set = {_up(r[0]) for r in _read_csv("gsi.csv")[1:]}

    eom = date(as_of.year, as_of.month, calendar.monthrange(as_of.year, as_of.month)[1])

    rows = []
    for r in data:
        g = lambda name: r[idx[name]]
        code = g("Cust Code")
        code_u = _up(code)
        val = g("Over Due Days")
        
        if isinstance(val, datetime):
            age = 0
        else:
            age = float(val or 0)
        ar = float(g("Ar Balance") or 0)
        val = ar if ar > 0 else 0.0
        b121 = val if age >= 121 else 0.0
        b120 = (val if age >= 91 else 0.0) - b121
        b90 = (val if age >= 61 else 0.0) - b120 - b121
        b60 = (val if age >= 31 else 0.0) - b90 - b120 - b121
        b30 = (val if age >= 16 else 0.0) - b60 - b90 - b120 - b121
        b15 = (val if age >= 0 else 0.0) - b30 - b60 - b90 - b120 - b121
        onacc = float(g("On Account") or 0)
        notdue = float(g("Not Due Amount") or 0)
        due = g("Document Due Date")
        due_d = due.date() if isinstance(due, datetime) else due if isinstance(due, date) else None
        adddue = notdue if (due_d and as_of < due_d <= eom) else 0.0
        brand = g("Brand")
        rows.append({
            "code": code, "name": g("Cust Name"), "main_ac": g("Main Ac"),
            "so_no": g("SO No"), "cust_region": _up(g("Cust Region")),
            "lpo": g("LPO No"), "doc": g("Document Number"),
            "docdt": g("Document Date"), "duedt": due,
            "ageing": g("Days From Docdt"), "overdue_days": g("Over Due Days"),
            "b15": b15, "b30": b30, "b60": b60, "b90": b90, "b120": b120,
            "b121": b121, "adddue": adddue,
            "bal": b15 + b30 + b60 + b90 + b120 + b121 + onacc + notdue,
            "onacc": onacc, "notdue": notdue, "brand": brand,
            "status": g("Customer Status"),
            "lcbg": float(g("LC & BG Guarantee") or 0),
            "auh": "AUH" if code_u in auh_set else "NOT AUH",
            "bum": bum_map.get(_up(brand), ""),
            "region": "KSA" if code_u.startswith("CK") else reg_map.get(_up(g("Cust Region")), ""),
            "shelly": "GSI" if code_u in gsi_set else "Not GSI",
            "renewals": "Renewals" if _up(g("Document Number")) in renew_set else "Not Renewals",
            "seafrica": sea_map.get(_up(g("Cust Region")), "NOT SE AFRICA"),
            "insurance": ins_map.get(code_u, 0.0),
            "backlog": bkl_map.get(code_u, 0.0),
            "pdc": pdc_map.get(code_u, 0.0),
        })

    return as_of, rows, ins_rows, bkl_map


def _group_cus(rows, by_brand: bool):
    """Group filtered rows Excel-pivot style; sorted by Overdue descending."""
    groups: dict = {}
    for r in rows:
        key = (r["code"], r["brand"]) if by_brand else r["code"]
        g = groups.get(key)
        if g is None:
            g = groups[key] = {
                "code": r["code"], "name": r["name"], "main_ac": r["main_ac"],
                "cust_region": r["cust_region"], "brand": r["brand"],
                "insurance": 0.0, "backlog": 0.0, "pdc": 0.0, "lcbg": 0.0,
                "onacc": 0.0, "notdue": 0.0, "adddue": 0.0,
                "b15": 0.0, "b30": 0.0, "b60": 0.0, "b90": 0.0,
                "b120": 0.0, "b121": 0.0,
            }
        for f in ("onacc", "notdue", "adddue", "b15", "b30", "b60", "b90", "b120", "b121"):
            g[f] += r[f]
        for f in ("insurance", "backlog", "pdc", "lcbg"):
            g[f] = max(g[f], r[f])
    out = list(groups.values())
    out.sort(key=lambda g: sum(g[b] for b in ("b15", "b30", "b60", "b90", "b120", "b121")), reverse=True)
    return out


def _cus_sheet(wb, fmts, name, as_of, groups, by_brand, lookup_cols):
    ws = wb.add_worksheet(name)
    heads = ["Code", "Cust Name", "Main Ac", "Region"]
    keys = ["code", "name", "main_ac", "cust_region"]
    if by_brand:
        heads.append("Brand")
        keys.append("brand")
    heads += [h for h, _ in lookup_cols] + [" On Account", " Not Due", "Additional due End of month"]
    keys += [k for _, k in lookup_cols] + ["onacc", "notdue", "adddue"]
    n0 = len(keys)  # 0-based index of ' Ar Balance'
    heads += [" Ar Balance", "Overdue", "Overdue + on account",
              "Ageing 1 to 15", "Ageing 16 to 30", "Ageing 31 to 60",
              "Ageing 61 to 90", "Ageing 91 to 120", "Ageing >=120"]
    col = xl_col_to_name
    L_bal, L_ovd, L_j = col(n0), col(n0 + 1), col(n0 - 3)  # ArBal, Overdue, OnAcc
    L_k = col(n0 - 2)  # Not Due
    L_p, L_u = col(n0 + 3), col(n0 + 8)  # first/last bucket
    first, last = 4, 3 + len(groups)

    ws.set_column(0, 0, 12)
    ws.set_column(1, 1, 40)
    ws.set_column(2, len(heads) - 1, 13, fmts["acc"])

    ws.write(0, 0, "AR Ageing Report as of ", fmts["bold"])
    ws.write_datetime(0, 2, as_of, fmts["date"])
    # ratio row + subtotal row
    for i in range(n0 - 3, len(heads)):
        c = col(i)
        ws.write_formula(1, i, f"=SUBTOTAL(9,{c}{first}:{c}{last})", fmts["acc_b"])
    ws.write_formula(0, n0 - 3, f"={L_j}2/{L_bal}2", fmts["pct"])
    ws.write_formula(0, n0 - 2, f"={L_k}2/{L_bal}2", fmts["pct"])
    ws.write_formula(0, n0 + 1, f"={L_ovd}2/{L_bal}2", fmts["pct"])
    for i in range(n0 + 3, n0 + 9):
        ws.write_formula(0, i, f"={col(i)}2/${L_ovd}$2", fmts["pct"])
    for i, h in enumerate(heads):
        ws.write(2, i, h, fmts["hdr"])
    for rn, g in enumerate(groups):
        er = first + rn
        for i, k in enumerate(keys):
            ws.write(er - 1, i, g[k])
        ws.write_formula(er - 1, n0, f"=SUM({L_j}{er},{L_k}{er},{L_ovd}{er})")
        ws.write_formula(er - 1, n0 + 1, f"=SUM({L_p}{er}:{L_u}{er})")
        ws.write_formula(er - 1, n0 + 2, f"=+{L_ovd}{er}+{L_j}{er}")
        for i, b in enumerate(("b15", "b30", "b60", "b90", "b120", "b121")):
            ws.write(er - 1, n0 + 3 + i, g[b])
    return L_bal


INV_HEADS = ["Cust Code", "Cust Name", "Main Ac", "SO No", "Txn Region", "LPO No",
             "Invoice", "Invoice Dt", "Invoice Due Dt", "Ageing", "Insured Credit",
             "Ageing 1 to 15", "Ageing 16 to 30", "Ageing 31 to 60", "Ageing 61 to 90",
             "Ageing 91 to 120", "Ageing >=120", "Additional due", "AR Balance",
             "On Account", "Not Due Amount", "Brand", "Customer Status"]
INV_KEYS = ["code", "name", "main_ac", "so_no", "cust_region", "lpo", "doc",
            "docdt", "duedt", "ageing", "overdue_days", "b15", "b30", "b60",
            "b90", "b120", "b121", "adddue", None, "onacc", "notdue", "brand", "status"]


def _inv_sheet(wb, fmts, name, rows, cus_name, cus_bal_col):
    ws = wb.add_worksheet(name)
    first, last = 4, 3 + len(rows)
    ws.set_column(0, 0, 12)
    ws.set_column(1, 1, 40)
    ws.set_column(2, 22, 13, fmts["acc"])
    ws.set_column(7, 8, 10, fmts["date"])
    ws.write_formula(0, 18, f"=+S2-'{cus_name}'!{cus_bal_col}2", fmts["acc_b"])
    for i in range(11, 21):  # L..U
        c = xl_col_to_name(i)
        ws.write_formula(1, i, f"=SUBTOTAL(9,{c}{first}:{c}{last})", fmts["acc_b"])
    for i, h in enumerate(INV_HEADS):
        ws.write(2, i, h, fmts["hdr"])
    for rn, r in enumerate(rows):
        er = first + rn
        for i, k in enumerate(INV_KEYS):
            if k is None:
                ws.write_formula(er - 1, i, f"=SUM(L{er}:Q{er},T{er}:U{er})")
            elif r[k] is not None:
                ws.write(er - 1, i, r[k])
    return ws


def build_bum_zip(*input_bytes) -> tuple[bytes, dict]:
    """Build all per-BUM files; returns (zip_bytes, meta)."""
    as_of, rows, ins_rows, bkl_map = compute_dataset(*input_bytes)
    cus_name = f"{as_of:%d.%m.%Y} CUS"
    inv_name = f"{as_of:%d.%m.%Y} Inv"

    def new_wb(buf):
        wb = xlsxwriter.Workbook(buf, {"default_date_format": _FMT_DATE})
        fmts = {
            "date": wb.add_format({"num_format": _FMT_DATE}),
            "hdr": wb.add_format({"bold": True, "font_color": "white", "bg_color": "#002060"}),
            "bold": wb.add_format({"bold": True}),
            "acc": wb.add_format({"num_format": _FMT_ACC}),
            "acc_b": wb.add_format({"num_format": _FMT_ACC, "bold": True}),
            "pct": wb.add_format({"num_format": "0.0%", "bold": True}),
        }
        return wb, fmts

    def cus_total(groups):
        return sum(
            g["onacc"] + g["notdue"] + g["b15"] + g["b30"] + g["b60"]
            + g["b90"] + g["b120"] + g["b121"]
            for g in groups
        )

    zbuf = io.BytesIO()
    checks = []  # per file: label, invoices, customers, inv total, cus total, match
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
        for label, pred, by_brand in FILE_SPECS:
            sel = [r for r in rows if pred(r)]
            groups = _group_cus(sel, by_brand)
            inv_tot = sum(r["bal"] for r in sel)
            cus_tot = cus_total(groups)
            checks.append({
                "file": f"AR - {label}", "invoices": len(sel),
                "customers": len(groups), "inv_total": inv_tot,
                "cus_total": cus_tot, "match": abs(inv_tot - cus_tot) < 0.01,
            })
            buf = io.BytesIO()
            wb, fmts = new_wb(buf)
            bal_col = _cus_sheet(wb, fmts, cus_name, as_of, groups,
                                 by_brand, CUS_LOOKUP_COLS)
            _inv_sheet(wb, fmts, inv_name, sel, cus_name, bal_col)
            wb.close()
            zf.writestr(f"AR - {label}.xlsx", buf.getvalue())

        # Ehab: everything, both CUS variants + insurance report
        groups = _group_cus(rows, False)
        inv_tot = sum(r["bal"] for r in rows)
        checks.append({
            "file": "Overall Region AR (Ehab)", "invoices": len(rows),
            "customers": len(groups), "inv_total": inv_tot,
            "cus_total": cus_total(groups),
            "match": abs(inv_tot - cus_total(groups)) < 0.01,
        })
        buf = io.BytesIO()
        wb, fmts = new_wb(buf)
        bal_col = _cus_sheet(wb, fmts, "Overall", as_of, groups,
                             False, EHAB_LOOKUP_COLS)
        _cus_sheet(wb, fmts, "Overall + Brand", as_of, _group_cus(rows, True),
                   True, EHAB_LOOKUP_COLS)
        _inv_sheet(wb, fmts, "Invoice Details", rows, "Overall", bal_col)
        ws = wb.add_worksheet("Insurance Report")
        ws.set_column(0, 0, 12)
        ws.set_column(1, 1, 45)
        ws.set_column(2, 6, 15, fmts["acc"])
        bal_by_code: dict[str, float] = {}
        for r in rows:
            bal_by_code[_up(r["code"])] = bal_by_code.get(_up(r["code"]), 0.0) + r["bal"]
        for i, h in enumerate(["Customer Code", "Customer name", "Country",
                               "Insurance Limit", "BACKLOG", "Net AR", "Available Limit"]):
            ws.write(0, i, h, fmts["hdr"])
        for rn, (code, name, country, limit) in enumerate(ins_rows, 1):
            ws.write(rn, 0, code)
            ws.write(rn, 1, name)
            ws.write(rn, 2, country)
            ws.write(rn, 3, limit)
            ws.write(rn, 4, bkl_map.get(_up(code), 0.0))
            ws.write(rn, 5, bal_by_code.get(_up(code), 0.0))
            ws.write_formula(rn, 6, f"=D{rn+1}-E{rn+1}-F{rn+1}")
        wb.close()
        zf.writestr("Overall Region AR (Ehab).xlsx", buf.getvalue())

    meta = {"as_of": as_of, "rows": len(rows), "checks": checks}
    return zbuf.getvalue(), meta
