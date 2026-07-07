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

**Key cells in the generated Excel:**
- **B5 — AR Data Date**: "the photo was taken on this date". Every formula compares against
  it. Months on/before B5 show 0 (that's the past, not a forecast). The first month after
  B5 is the "active month".
- **Columns K–V**: the aging buckets per customer (data, not formulas).
- **Columns L–P**: the "Not Due" money split by *when* it will become due (0-30 days from
  the AR date, 31-60, 61-90, 91-180, 180+). Needed because "not due" money also starts
  aging once its due date passes.
- **Orange columns** (per month): input cells. "Collections FC (FIFO)" = expected payment,
  auto-allocated oldest-bucket-first. "Specific Alloc" = same thing but the manager says
  exactly which bucket the payment belongs to. Use one or the other, not both.
- **Per month, 3 formula columns**: *Expected AR* (balance after collections), *AR
  Provision FC* (forecasted provision for that month-end), *Provision Effect* (change vs
  the previous month — what hits the P&L).
- **Row 9**: totals (SUBTOTAL formulas, so they follow any filter you apply).

---

## 2. What the tool automates

Before: someone built this workbook by hand every month (copy data in, fix formulas...).

Now (third tab in the Streamlit app):

```
inputs:  tool-1 output workbook (By_Customer + Invoice sheets)
         AR Data Date  (date picker)
         Insurance Master (optional Excel)
output:  AR Collection and Provision Forecast - <Month>.xlsx
         (single "ALL" sheet, all formulas live, ready for credit managers)
```

The user flow: pick date → upload file(s) → click download. ~15–20 s for ~1,700 customers.

---

## 3. How the code works (`provision/` package)

```
ui.py      Streamlit tab. Reads the upload, shows warnings, download button.
mapper.py  pandas: By_Customer sheet -> one row per customer with the
           fixed-column values (A-Y). Also:
           - insurance lookup (same logic as the BUD2026 tool)
           - Not Due breakdown L-P: buckets each invoice's Document Due Date
             vs the AR Data Date (invoice sheet found automatically:
             "Invoice", "AR_Backlog" or "Traverse_AR")
           - infer_as_on_date(): recovers the file's real as-on date
             (Document Date + Ageing days) to warn on mismatch
export.py  xlsxwriter: writes the ALL sheet - titles, B5, rates row,
           SUBTOTAL totals, headers, data rows, and one formula per
           formula-column per row, taken from template_data.json.
template_data.json  THE IMPORTANT FILE. Generated, not hand-written.
           Contains the exact formulas, headers, number formats, column
           widths and colors extracted from the original .xlsb model.
```

### Where the formulas came from (and the 8192 story)

The original model is an `.xlsb` (binary Excel). Python libraries **cannot read formulas
from .xlsb** — so the formulas were extracted with Excel itself (PowerShell COM automation:
open the file, read every cell's `.Formula`, dump to JSON), then turned into row-templates
(`B12` → `B«R»`, and `«R»` is replaced with the real row number at write time).

The formulas use Excel-365 `LET(...)` functions. Inside an `.xlsx` file, every LET variable
is silently stored with a `_xlpm.` prefix (and `LET` as `_xlfn.LET`). Two consequences:

1. `export.py::_to_stored_form()` adds those prefixes when writing — without them Excel
   shows `#NAME?`.
2. The stored text of the Oct/Nov/Dec "AR Provision FC" formulas exceeds the **8,192-char
   limit of the .xlsx format** (that's why Excel refuses to save the original as .xlsx).
   Fix: in `template_data.json` those three formulas (columns EJ, EU, FF) have mechanically
   shortened LET variable names (`cumSpecQ` → `sq` etc.). Same logic, same results, fits.

### How we know it's correct

A verification run rebuilt the workbook from the original file's own data, then Excel
(via COM) fully recalculated both files and compared **all 40 formula columns × 1,714
customers: zero differences**, totals equal, and typing a test collection moved both
files identically.

### Maintenance notes

- **FY 2026 is hardcoded** in every monthly formula (dates like `DATE(2026,6,1)`), same as
  the original model. For 2027 you must regenerate `template_data.json` from a 2027 model
  (COM-dump the formulas again) — or generalize the year in the templates.
- Columns **X/Y** ("AR Provision at \<prior date\>") are left **blank on purpose** — finance
  fills them manually. Until then, column AC (= AB − X) just equals AB.
- If the uploaded workbook has **no invoice-level sheet**, L–P can't be computed — the tool
  puts the whole Not Due amount into column L and shows a warning.
- The dummy sample file in this repo has an **active filter** — its row-9 totals only count
  visible rows. Not a bug; SUBTOTAL is designed to do that.
