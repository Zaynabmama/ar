# Sales Report tool

Builds the quarterly Sales Report output sheets from the raw sales export.

```
python -m sales_report.logic "Sales Report Q2-2026.xlsx" "Sales Report output.xlsx" 2026Q2
```

The last argument keeps only invoices whose Doc Dt is in that quarter (the margin report is cumulative). Leave it out to use every row.

Input: the export workbook; the sheet named `Data` (or the first sheet) with the columns
Invoice No, Pay Terms, Cust Code, Cust Name, Customer Region, Adjusted Sales.

How invoices are sorted:
1. Customer code with 4, 5 or 6 as the 3rd character -> **Aff**.
2. Otherwise by pay term, from `sales_report/data/pay_terms.csv` (confirmed with credit).
3. A pay term not in the CSV goes to **Open** and is listed on a "New pay terms" sheet.
   Add it to the CSV with the right bucket and run again.

`Basic Data` (first sheet) has one row per customer with the total of each payment-term group in its own column, plus SR. Region, Insured limit, By DCL, Check and diff are left empty for the later manual steps.

`Inv By SR` lists every credit note (SR / SRCL) with a non-zero amount; those invoices also stay in their bucket.
Check cells: `XXX-By Inv`!K1 and `By Cus`!G1 must be 0.

Streamlit tab: "Sales by Payment Terms" (sales_report/ui.py).

Tested against Sales Report Q2-2026: all 20 output sheets match the manual file.
