# AR Provision Forecast Tool — Plain-Language Guide

*For someone with zero finance background. First the business story, then how the code works.*

---

## 1. The business story (finance from zero)

**AR (Accounts Receivable)** = money customers owe us. We sold them something, sent an
invoice, and they haven't paid yet.

**The problem:** some customers will never pay. Accounting rules say we must estimate that
expected loss *in advance* and record it as an expense. That estimate is called the
**provision** (also "provision for bad debts").

**How do you estimate it?** By age. The longer an invoice stays unpaid past its due date,
the less likely it ever gets paid. So all invoices are sorted into **aging buckets**:

| Bucket | Meaning | Provision rate |
|---|---|---|
| Not Due | due date hasn't arrived yet | 3% |
| 1–30 | 1–30 days late | 3% |
| 31–60 | 31–60 days late | 3% |
| 61–90 | 61–90 days late | **25%** |
| 91–120 | | **50%** |
| 121–150 | | **75%** |
| ≥151 | 151+ days late | **100%** (we assume it's lost) |

Provision = amount in each bucket × that bucket's rate, summed up.

**The forecast part:** time moves. An invoice that is 40 days late today will be 70 days
late next month — it *shifts one bucket forward every month*, and its rate jumps (25% →
50%...). So even if nothing else changes, **the provision grows every month by itself**.
This file forecasts that growth for every month of 2026.

**Collections fight the growth.** A credit manager can say "I expect customer X to pay
500,000 in August". That payment removes money from the buckets — the model removes it
from the **oldest (riskiest) bucket first** (that's the **FIFO** rule) — so the forecasted
provision goes down. That's the whole game of the file: *type expected collections, watch
the provision forecast react.*

**Two refinements:**
- **Insurance**: some customers are credit-insured. If they don't pay, the insurer pays ~95%.
  So the insured part of their balance is provisioned at only 5% of the normal rate.
- **Only account 12301 is provisioned** — other account types (12305 etc.) are shown but
  always get provision 0. That's a business rule from finance.

**Key cells in the generated Excel** (new "Master File" layout, July 2026):
- **B3 — AR Data Date**: "the photo was taken on this date". Every formula compares against
  it. Months on/before B3 show 0 (that's the past, not a forecast). The first month after
  B3 is the "active month".
- **Columns J–U**: Not Due total (J) + the aging buckets per customer (data, not formulas).
- **Columns K–O**: the "Not Due" money split by *when* it will become due (0-30 days from
  the AR date, 31-60, 61-90, 91-180, 180+). Needed because "not due" money also starts
  aging once its due date passes. Taken directly from the By_Customer "Not Due ..."
  columns (the collectible view produced by the AR Backlog tool) — the master file's
  own K–O values follow the same status-filtered rule.
- **Column V — AR Balance**: a live formula (= buckets + Not Due + On Account).
- **Input columns** (per month): "Collections FC (FIFO)" = expected payment,
  auto-allocated oldest-bucket-first. "Specific Alloc" = same thing but the manager says
  exactly which bucket the payment belongs to. Use one or the other, not both.
- **Per month, 3 forecast formula columns**: *Expected AR* (balance after collections), *AR
  Provision FC* (forecasted provision for that month-end, computed by the workbook-level
  `MW_PROV_FC_CORE` function), *Provision Effect* (Provision FC − column AA, the base
  provision — what hits the P&L).
- **Per month, 4 "Actuals" columns** (new): type the real ERP collection into *Actual
  Collection* and the three formula columns light up — *Expected AR − Actual*, *Actual AR
  Provision* (same logic, actual collections instead of forecast), and *Variance Actual vs
  Forecast*. While Actual Collection is blank, they stay blank.
- **Columns HF–IC** (far right): cumulative-deduction helper columns used by Expected AR.
- **Row 7**: totals (SUBTOTAL formulas, so they follow any filter you apply).

---

## 2. What the tool automates

Before: someone built this workbook by hand every month (copy data in, fix formulas...).

Now (third tab in the Streamlit app):

```
inputs:  tool-1 output workbook (only the By_Customer sheet is read)
         AR Data Date  (date picker)
         Insurance Master (optional Excel)
output:  AR Collection and Provision Forecast - <Month>.xlsx
         (single "ALL" sheet, all formulas live, ready for credit managers)
```

The user flow: pick date → upload file(s) → click download. Roughly a minute for ~5,500
customers (the new model writes ~100 formula columns per customer).

---

## 3. How the code works (`provision/` package)

```
ui.py      Streamlit tab. Reads the upload, shows warnings, download button.
mapper.py  pandas: By_Customer sheet -> one row per customer with the
           fixed-column values (A-U). Also:
           - insurance lookup (same logic as the BUD2026 tool)
           - Not Due breakdown K-O: read straight from the By_Customer
             "Not Due 0-30/31-60/61-90/91-180/180+" columns that the AR
             Backlog tool now produces (no invoice-level sheet is opened)
export.py  xlsxwriter: writes the ALL sheet - titles, B3, rates row (5),
           SUBTOTAL totals (row 7), headers (row 9), data rows (10+), one
           formula per formula-column per row, and the MW_PROV_FC_CORE
           LAMBDA as a workbook defined name - all from template_data.json.
template_data.json  THE IMPORTANT FILE. Generated, not hand-written.
           Contains the exact formulas, the LAMBDA text, headers, number
           formats, column widths and colors extracted from
           "AR Collection and Provision Forecast - Master File.xlsb".
```

### Where the formulas came from (and the 8192 story)

The model is an `.xlsb` (binary Excel). Python libraries **cannot read formulas
from .xlsb** — so the formulas were extracted with Excel itself (PowerShell COM automation:
open the file, read every cell's `.Formula`, dump to text), then turned into row-templates
(`B10` → `B«R»`, and `«R»` is replaced with the real row number at write time).

The heavy provision math lives in **`MW_PROV_FC_CORE`**, a workbook-level `LAMBDA` function
(defined name) that the 24 monthly provision columns call with cumulative collections as
arguments. Inside an `.xlsx` file, Excel silently stores every LAMBDA/LET parameter with a
`_xlpm.` prefix (and `LET`/`LAMBDA` as `_xlfn.LET`/`_xlfn.LAMBDA`).
`export.py::_to_stored_form()` adds those prefixes when writing — without them Excel shows
`#NAME?`. Because the shared logic sits in the LAMBDA (stored once), the per-cell formulas
stay far below the 8,192-char stored-formula limit that plagued the previous model.

### How we know it's correct

A verification run rebuilt the workbook from the master file's own customer data, pasted
its 12,790 non-zero input cells, let Excel (via COM) fully recalculate, and compared
**1,296,390 cells — all 101 formula columns × 5,470 customers plus 200 SUBTOTALs: zero
differences**. A separate end-to-end run exercised mapper + export on a synthetic
By_Customer file (insurance netting, non-12301 gating and the Actual-columns blank
behavior all hand-checked).

### Maintenance notes

- **FY 2026 is hardcoded** in every monthly formula (dates like `DATE(2026,6,1)`), same as
  the master model. For 2027 you must regenerate `template_data.json` from a 2027 model
  (COM-dump the formulas again) — or generalize the year in the templates.
- Columns **W/X** ("AR Provision at ...") are left **blank on purpose** — finance fills
  them manually. Until then, column AB (= AA − W) just equals AA. Column AC (Notes) is
  also a manual column.
- **Actual Collection columns are blank on purpose** (not 0): the three Actual formula
  columns only calculate once a value is typed. The export must never write 0 there.
- The helper columns HF–IC gate months against `$B{row-7}` — for the first data row that is
  `$B3` (the AR Data Date), for every later row an empty cell (gate always true). That is a
  fill-down artifact **in the master itself**, replicated deliberately (user decision,
  2026-07-13). It only matters if someone types collections for months on/before the AR
  Data Date.
- Some labels are copied verbatim from the master even where stale (e.g. I6 says
  "Balance at 31-03-2026" regardless of the chosen AR date) — user decision, 2026-07-13.
- If the uploaded By_Customer sheet has **no "Not Due ..." breakdown columns** (a file made
  with an older AR Backlog version), the tool puts the whole Not Due amount into column K
  and shows a warning asking to regenerate the file.
- Since the breakdown is the **collectible view**, 2027+ dues are included in "180+" (the
  master excluded them) and non-collectible statuses (DOUBTFUL etc.) show zeros — matching
  the master's own K–O values, which follow the same status rule.
