"""Sales Report builder.

Takes the raw sales export (sheet "Data") and rebuilds every output sheet of
the quarterly Sales Report:

    Data -> one row per invoice -> one bucket per invoice -> totals per customer

Bucket rules (confirmed with credit):
  1. Aff        customer code with 4, 5 or 6 as 3rd character
  2. by pay term, from data/pay_terms.csv  (Fin CN, LC, Adv, Above MCP, SCF, Open)
  3. a pay term that is not in the CSV goes to Open and is listed on the
     "New pay terms" sheet so it can be added to the CSV
Inv By SR is an extra view of all credit notes (invoice number starting SR).
"""
from __future__ import annotations

import csv
import io
import re
import zipfile
from datetime import datetime
from collections import OrderedDict
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

DATA_DIR = Path(__file__).parent / "data"
PAY_TERMS_CSV = DATA_DIR / "pay_terms.csv"

AFFILIATE_CHARS = {"4", "5", "6"}
BLANK_TERM = "(blank)"

# Bucket -> (invoice sheet, customer sheet), in the order the check formulas add them.
BUCKETS = OrderedDict([
    ("Fin CN", ("Inv By Fin CN", "--By FN CN")),
    ("Aff", ("Inv By Aff", "--By Aff")),
    ("LC", ("Inv By LC", "--By LC")),
    ("Adv", ("Inv By Adv", "--By Adv")),
    ("Above MCP", ("Inv Above MCP", "--By Above MCP")),
    ("SCF", ("Inv By SCF", "--By SCF")),
    ("Open", ("Inv By Open", "---By Open")),
])
SR_SHEETS = ("Inv By SR", "--By SR")

REQUIRED = ["Invoice No", "Pay Terms", "Cust Code", "Cust Name",
            "Customer Region", "Adjusted Sales"]

INV_HEADERS = ["Invoice No", "Pay Terms", "Cust Code", "Cust Name",
               "Customer Region", "Total"]
CUS_HEADERS = ["Cust Code", "Customer Region", "Total"]
NUM_FMT = "#,##0.00"
BOLD = Font(bold=True)
HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")


# ---------------------------------------------------------------- loading
def _repair(raw: bytes) -> io.BytesIO:
    """The ERP export has an out-of-range font family in styles.xml, which
    openpyxl refuses. Rewrite that one value so the file opens."""
    src = zipfile.ZipFile(io.BytesIO(raw))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "xl/styles.xml":
                data = re.sub(
                    rb'<family val="(\d+)"/>',
                    lambda m: m.group(0) if int(m.group(1)) <= 14 else b'<family val="2"/>',
                    data,
                )
            dst.writestr(item, data)
    out.seek(0)
    return out


def load_data(file) -> list[dict]:
    """Read the Data sheet (or the first sheet) and return the needed columns."""
    raw = file.read() if hasattr(file, "read") else Path(file).read_bytes()
    wb = openpyxl.load_workbook(_repair(raw), read_only=True, data_only=True)
    ws = wb["Data"] if "Data" in wb.sheetnames else wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    header = [str(h).strip() if h is not None else "" for h in next(rows)]
    missing = [c for c in REQUIRED if c not in header]
    if missing:
        raise ValueError(f"Data sheet is missing columns: {', '.join(missing)}")
    idx = {c: header.index(c) for c in REQUIRED}
    out = []
    for r in rows:
        inv = r[idx["Invoice No"]]
        if inv in (None, ""):
            continue
        rec = {c: r[idx[c]] for c in REQUIRED}
        rec["Adjusted Sales"] = float(rec["Adjusted Sales"] or 0)
        for c in ("Pay Terms", "Cust Code", "Cust Name", "Customer Region"):
            rec[c] = "" if rec[c] is None else str(rec[c])
        rec["Invoice No"] = str(inv)
        if "Doc Dt" in header:
            rec["Doc Dt"] = r[header.index("Doc Dt")]
        out.append(rec)
    wb.close()
    return out


def filter_quarter(lines: list[dict], year: int, quarter: int) -> list[dict]:
    """Keep only the lines whose Doc Dt falls in the quarter (the margin report is cumulative)."""
    months = range(3 * quarter - 2, 3 * quarter + 1)
    kept = []
    for r in lines:
        d = r.get("Doc Dt")
        if d is None:
            raise ValueError("Data sheet has no Doc Dt column; cannot filter by quarter")
        if isinstance(d, str):
            d = datetime.strptime(d[:10], "%Y-%m-%d")
        if d.year == year and d.month in months:
            kept.append(r)
    return kept


def load_pay_terms(path: Path = PAY_TERMS_CSV) -> dict[str, str]:
    with open(path, newline="", encoding="utf-8") as f:
        return {r["pay_term"]: r["bucket"] for r in csv.DictReader(f)}


# ---------------------------------------------------------------- logic
def by_invoice(lines: list[dict]) -> list[list]:
    """One row per invoice, in order of first appearance (like the pivot)."""
    # Excel pivots group text ignoring upper/lower case and show the first spelling seen.
    # The spelling is fixed per column across the whole sheet, like the pivot's item list.
    spelling = [{} for _ in range(5)]
    inv: OrderedDict[tuple, list] = OrderedDict()
    for r in lines:
        row = [r["Invoice No"], r["Pay Terms"], r["Cust Code"], r["Cust Name"], r["Customer Region"]]
        row = [spelling[i].setdefault(v.casefold(), v) for i, v in enumerate(row)]
        key = tuple(v.casefold() for v in row)
        if key not in inv:
            inv[key] = row + [0.0]
        inv[key][5] += r["Adjusted Sales"]
    return list(inv.values())


def is_affiliate(cust_code: str) -> bool:
    return len(cust_code) >= 3 and cust_code[2] in AFFILIATE_CHARS


def bucket_of(row: list, terms: dict[str, str]) -> tuple[str, bool]:
    """Return (bucket, known_term)."""
    if is_affiliate(row[2]):
        return "Aff", True
    term = row[1] or BLANK_TERM
    if term in terms:
        return terms[term], True
    return "Open", False


def by_customer(rows: list[list]) -> list[list]:
    agg: OrderedDict[tuple, float] = OrderedDict()
    for r in rows:
        key = (r[2], r[4])
        agg[key] = agg.get(key, 0.0) + r[5]
    return [[k[0], k[1], v] for k, v in agg.items()]


def build(lines: list[dict], terms: dict[str, str] | None = None) -> dict:
    terms = terms if terms is not None else load_pay_terms()
    invoices = by_invoice(lines)
    split = {b: [] for b in BUCKETS}
    new_terms: OrderedDict[str, list] = OrderedDict()
    for row in invoices:
        b, known = bucket_of(row, terms)
        split[b].append(row)
        if not known:
            t = new_terms.setdefault(row[1] or BLANK_TERM, [0, 0.0])
            t[0] += 1
            t[1] += row[5]
    # Credit notes with a non-zero amount (a fully reversed SR nets to 0 and is left out).
    sr = [r for r in invoices if r[0].upper().startswith("SR") and abs(r[5]) > 0.005]
    return {"invoices": invoices, "split": split, "sr": sr, "new_terms": new_terms,
            "total": sum(r[5] for r in invoices)}


# ---------------------------------------------------------------- writing
def _head(ws, row, headers):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = BOLD
        c.fill = HEAD_FILL


def _widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _invoice_sheet(wb, name, rows, helper=True):
    ws = wb.create_sheet(name)
    last = max(len(rows) + 2, 3)
    ws["F1"] = f"=SUBTOTAL(9,F3:F{last})"
    ws["F1"].number_format = NUM_FMT
    ws["F1"].font = BOLD
    _head(ws, 2, INV_HEADERS)
    if helper:
        ws["H2"] = "Cust Code"
        ws["H2"].font = BOLD
    for i, r in enumerate(rows, 3):
        for j, v in enumerate(r, 1):
            ws.cell(row=i, column=j, value=v)
        ws.cell(row=i, column=6).number_format = NUM_FMT
        if helper:
            ws.cell(row=i, column=8, value=f"=MID(C{i},3,1)")
    ws.auto_filter.ref = f"A2:{'H' if helper else 'F'}{last}"
    ws.freeze_panes = "A3"
    _widths(ws, [16, 34, 12, 45, 16, 16, 3, 10])
    return ws


def _customer_sheet(wb, name, rows, subtotal=False):
    ws = wb.create_sheet(name)
    last = max(len(rows) + 2, 3)
    ws["C1"] = f"=SUBTOTAL(9,C3:C{last})" if subtotal else f"=SUM(C3:C{last})"
    ws["C1"].number_format = NUM_FMT
    ws["C1"].font = BOLD
    _head(ws, 2, CUS_HEADERS)
    for i, r in enumerate(rows, 3):
        for j, v in enumerate(r, 1):
            ws.cell(row=i, column=j, value=v)
        ws.cell(row=i, column=3).number_format = NUM_FMT
    ws.freeze_panes = "A3"
    _widths(ws, [14, 18, 18])
    return ws


def _pivot_style(wb, name, title, headers, rows, total_col):
    """Pivot results written as values, laid out like the Excel pivot."""
    ws = wb.create_sheet(name)
    ws["A3"] = title
    ws["A3"].font = BOLD
    _head(ws, 4, headers)
    for i, r in enumerate(rows, 5):
        for j, v in enumerate(r, 1):
            ws.cell(row=i, column=j, value=v)
        ws.cell(row=i, column=total_col).number_format = NUM_FMT
    g = len(rows) + 5
    ws.cell(row=g, column=1, value="Grand Total").font = BOLD
    c = ws.cell(row=g, column=total_col, value=sum(r[total_col - 1] for r in rows))
    c.number_format = NUM_FMT
    c.font = BOLD
    ws.freeze_panes = "A5"
    return ws


BASIC_HEADERS = ["Cust Code", "Name", "Customer Country", "Region", "Insured limit", "Total",
                 "Open Credit", "LC", "Finance CN", "Advance ", "Affiliates", "SCF",
                 "Above MCP", "By DCL", "Total", "SR", "Check", "diff"]
# Basic Data column -> bucket whose per-customer total fills it.
BASIC_BUCKETS = ["Open", "LC", "Fin CN", "Adv", "Aff", "SCF", "Above MCP"]


def basic_data(result: dict) -> list[list]:
    """One row per customer (same rows and order as By Cus) with the total of each
    payment-term group in its own column, plus the credit notes (SR)."""
    names: dict[str, str] = {}
    for r in result["invoices"]:
        names.setdefault(r[2], r[3])
    per = {b: {} for b in BASIC_BUCKETS}
    for b in BASIC_BUCKETS:
        for code, _region, total in by_customer(result["split"][b]):
            per[b][code] = per[b].get(code, 0.0) + total
    sr: dict[str, float] = {}
    for code, _region, total in by_customer(result["sr"]):
        sr[code] = sr.get(code, 0.0) + total
    rows = []
    for code, region, total in by_customer(result["invoices"]):
        rows.append([code, names.get(code, ""), region, None, None, total]
                    + [per[b].get(code, 0) for b in BASIC_BUCKETS]
                    + [None, None, sr.get(code, 0), None, None])
    return rows


def _basic_data_sheet(wb, result: dict):
    """'Basic Data': the per-customer table used for the insurance / ITO checks.
    Region, Insured limit, By DCL, Check and diff are left empty for the later manual steps."""
    ws = wb.create_sheet("Basic Data", 0)
    rows = basic_data(result)
    first, last = 8, max(len(rows) + 7, 8)
    ws["A1"] = "To check Sales from CD account to Oman and Kuwait"
    ws["A2"] = "To check duplicate accounts in Orion codes"
    ws["L3"] = "Diff Check"
    ws["M3"] = "=O5-F5"
    ws["M3"].number_format = NUM_FMT
    for col in "FGHIJKLMNOP":
        c = ws[f"{col}5"]
        c.value = f"=SUBTOTAL(9,{col}{first}:{col}{last})"
        c.number_format = NUM_FMT
        c.font = BOLD
    _head(ws, 7, BASIC_HEADERS)
    for i, r in enumerate(rows, first):
        for j, v in enumerate(r, 1):
            if v is not None:
                ws.cell(row=i, column=j, value=v)
        ws.cell(row=i, column=15, value=f"=SUM(G{i}:N{i})")
        for j in range(6, 17):
            ws.cell(row=i, column=j).number_format = NUM_FMT
    ws.auto_filter.ref = f"A7:R{last}"
    ws.freeze_panes = "C8"
    _widths(ws, [12, 40, 16, 10, 12] + [15] * 11 + [10, 10])
    return ws


def write_basic_data(result: dict, out) -> None:
    """Second output: a workbook with only the Basic Data sheet."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _basic_data_sheet(wb, result)
    wb.save(out)


def write(result: dict, out) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    split, sr, invoices = result["split"], result["sr"], result["invoices"]

    p = _pivot_style(wb, "Piv By Inv", "Sum of Adjusted Sales", INV_HEADERS, invoices, 6)
    _widths(p, [16, 34, 12, 45, 16, 16])
    p = _pivot_style(wb, "XXX Piv By Cus", "Sum of Total", CUS_HEADERS, by_customer(sr), 3)
    _widths(p, [14, 18, 18])

    # By Cus: all invoices per customer, checked against the bucket sheets.
    ws = _customer_sheet(wb, "By Cus", by_customer(invoices), subtotal=True)
    ws["E1"] = "=" + "+".join(f"'{cus}'!C1" for _, cus in BUCKETS.values())
    ws["G1"] = "=C1-E1"
    for c in ("E1", "G1"):
        ws[c].number_format = NUM_FMT
        ws[c].font = BOLD

    ws = _invoice_sheet(wb, "XXX-By Inv", invoices)
    ws["J1"] = "=" + "+".join(f"'{inv}'!F1" for inv, _ in BUCKETS.values())
    ws["K1"] = "=J1-F1"
    for c in ("J1", "K1"):
        ws[c].number_format = NUM_FMT
        ws[c].font = BOLD

    for b, (inv, _) in BUCKETS.items():
        _invoice_sheet(wb, inv, split[b])
    _invoice_sheet(wb, SR_SHEETS[0], sr, helper=False)

    for b, (_, cus) in BUCKETS.items():
        _customer_sheet(wb, cus, by_customer(split[b]), subtotal=b in ("Adv", "Above MCP"))
    _customer_sheet(wb, SR_SHEETS[1], by_customer(sr))

    if result["new_terms"]:
        ws = wb.create_sheet("New pay terms", 0)
        ws["A1"] = ("Pay terms not in data/pay_terms.csv. They were put in Open. "
                    "Add them to the CSV with the right bucket and run again.")
        ws["A1"].font = Font(bold=True, color="C00000")
        _head(ws, 3, ["Pay term", "Invoices", "Amount", "Put in"])
        for i, (t, (n, s)) in enumerate(result["new_terms"].items(), 4):
            ws.cell(row=i, column=1, value=t)
            ws.cell(row=i, column=2, value=n)
            ws.cell(row=i, column=3, value=s).number_format = NUM_FMT
            ws.cell(row=i, column=4, value="Open")
        _widths(ws, [60, 10, 18, 10])

    wb.save(out)


def run(data_file, out, year: int | None = None, quarter: int | None = None,
        basic_out=None) -> dict:
    """Read the export, keep the quarter (if given), build, write. Returns a summary."""
    lines = load_data(data_file)
    if year and quarter:
        lines = filter_quarter(lines, year, quarter)
    result = build(lines)
    write(result, out)
    if basic_out is not None:
        write_basic_data(result, basic_out)
    return summary(result)


def summary(result: dict) -> dict:
    s = {b: (len(v), sum(r[5] for r in v)) for b, v in result["split"].items()}
    s["SR (view)"] = (len(result["sr"]), sum(r[5] for r in result["sr"]))
    s["All invoices"] = (len(result["invoices"]), result["total"])
    s["New pay terms"] = len(result["new_terms"])
    return s


if __name__ == "__main__":
    import sys
    if len(sys.argv) not in (3, 4):
        sys.exit("usage: python -m sales_report.logic <export.xlsx> <output.xlsx> [2026Q2]")
    yq = sys.argv[3].upper().split("Q") if len(sys.argv) == 4 else (None, None)
    year, quarter = (int(yq[0]), int(yq[1])) if yq[0] else (None, None)
    basic = str(Path(sys.argv[2]).with_name(Path(sys.argv[2]).stem + ' - Basic Data.xlsx'))
    for k, v in run(sys.argv[1], sys.argv[2], year, quarter, basic_out=basic).items():
        print(f"{k:14} {v}")
