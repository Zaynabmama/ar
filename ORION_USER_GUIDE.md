# Orion Tool - AR Backlog → By_Customer Forecast Tool
## User Guide & System Documentation

---

## Overview
The **Orion Tool** is an AR (Accounts Receivable) forecasting system that processes AR Backlog data and generates forward-looking collections forecasts grouped by customer. It takes invoice-level detail and produces a customer-level summary with period-based (quarterly) collection projections.

**Purpose**: Transform raw AR backlog into actionable customer-level forecasts with aging analysis and multi-quarter projections.

---

## Step-by-Step User Workflow

### Step 1: Access the Tool
- Navigate to the "AR Backlog" tab in the application
- Select **"Orion"** as your data source

### Step 2: Choose Starting Quarter
- Use the **"Starting Quarter"** dropdown
- Available options: **Q1, Q2, Q3, Q4** (all based on 2026)
- **Why it matters**: Determines which quarters/periods will appear in the output and how invoices are grouped by due date

### Step 3: Upload AR Backlog Excel File
- Click **"Upload AR Backlog Excel"**
- File must be in **.xlsx or .xls** format
- **Critical requirements**:
  - Cells must contain an "As on Date" (evaluation date) in **cell B14**
  - Data header row must be at **row 16**
  - All data starts from **row 17 onwards**

### Step 4: System Processing
Once you upload the file, the tool automatically:
1. **Reads** the "As on Date" from cell B14 (the reference date for aging calculations)
2. **Extracts** all invoice rows starting from row 17
3. **Calculates** aging days, overdue days, and categorizes each invoice
4. **Groups** invoices by customer and main account
5. **Allocates** amounts to future quarters based on due dates
6. **Generates** forecasts and formulas for collection planning

### Step 5: Review Output Sheets
The tool generates an Excel file with **3 sheets**:

#### **Sheet 1: AR_Backlog** (Detail)
- Shows all original invoice records plus derived columns
- One row = One invoice
- Use this to validate source data and review aging details

#### **Sheet 2: By_Customer** (Summary + Forecast)
- One row = One customer/Main Account combination
- Aggregated amounts by aging bracket and by quarter/period
- **Includes Excel formulas** for manual entry of collection actuals and forecast updates
- This is the **primary working sheet** for forecasting

#### **Sheet 3: Invoice** (Detailed Invoice List)
- Filtered and formatted invoice-level view
- Includes customer, document details, aging, and amounts by bucket
- For detailed invoice review and drill-down

### Step 6: Download & Use
- Click **"Download Processed File"**
- File name: **processed_AR_backlog.xlsx**
- Open in Excel and use the **By_Customer** sheet for forecasting

---

## Input File Requirements

### Mandatory Columns in Your Excel
| Column Name | Purpose | Format | Example |
|---|---|---|---|
| **Cust Code** | Customer identifier | Text/Number | CP001 |
| **Cust Name** | Customer name | Text | ABC Company |
| **Main Ac** | Main account code | Text | 11001 |
| **Cust Region** | Customer region (optional, used for classification) | Text | KSA, UAE |
| **Document Number** | Invoice number | Text | INV-2026-001 |
| **Document Date** | Invoice issue date | Date | 2026-01-15 |
| **Document Due Date** | Payment due date | Date | 2026-02-15 |
| **Ar Balance** | Outstanding invoice amount | Number | 50000 |
| **Customer Status** | Payment status (optional) | Text | GOOD, REGULAR, SUBSTANDARD |
| **Payment Terms** | Terms of payment | Text | Net 30 |
| **Brand** | Product/service brand | Text | Brand A |
| **Total Insurance Limit** | Customer insurance limit | Number | 100000 |
| **LC & BG Guarantee** | LC/BG amount | Number | 25000 |
| **SO No** | Sales Order number | Text | SO-2026-001 |
| **LPO No** | Purchase Order number | Text | LPO-2026-001 |

**As on Date (Cell B14)**: The reference date used for all aging calculations. This is critical—all aging is measured from this date.

---

## Output Columns - AR_Backlog Sheet

### Core Input Columns (Preserved as-is)
All original columns from your input file are retained for reference.

### Derived Columns (Added by System)

| Column | Source | Calculation | Purpose |
|---|---|---|---|
| **Ageing (Days)** | Document Date | As on Date - Document Date | Days since invoice issued |
| **Overdue days (Days)** | Document Due Date | As on Date - Document Due Date | Days past due (0+ means overdue) |
| **Region (Derived)** | Cust Region + Cust Code | Rule-based classification logic | Standardized region field |
| **Ar Balance (Copy)** | Ar Balance | Copy of original | Preserved for calculations |
| **Aging Bracket (Label)** | Overdue days | Categorized label (see below) | Text label for aging bucket |
| **Updated Status** | Customer Status | Filled with SUBSTANDARD if empty | Standardized customer status |
| **Invoice Value (Derived)** | Ar Balance | MAX(Ar Balance, 0) | Positive balance only |
| **On Account (Derived)** | Ar Balance | MIN(Ar Balance, 0) as absolute value | Credit/on-account amounts |
| **Not Due (Derived)** | Overdue days + Ar Balance | Amount if Overdue days ≤ 0 | Amounts not yet due |
| **Aging 1 to 30 (Amount)** | Overdue days + Ar Balance | Amount if 0 < Overdue days ≤ 30 | Amount in 1-30 days overdue bracket |
| **Aging 31 to 60 (Amount)** | Overdue days + Ar Balance | Amount if 30 < Overdue days ≤ 60 | Amount in 31-60 days overdue bracket |
| **Aging 61 to 90 (Amount)** | Overdue days + Ar Balance | Amount if 60 < Overdue days ≤ 90 | Amount in 61-90 days overdue bracket |
| **Aging 91 to 120 (Amount)** | Overdue days + Ar Balance | Amount if 90 < Overdue days ≤ 120 | Amount in 91-120 days overdue bracket |
| **Aging 121 to 150 (Amount)** | Overdue days + Ar Balance | Amount if 120 < Overdue days ≤ 150 | Amount in 121-150 days overdue bracket |
| **Aging >=151 (Amount)** | Overdue days + Ar Balance | Amount if Overdue days > 150 | Amount in 151+ days overdue bracket |
| **Ageing > 365 (Amt)** | Ageing days + Ar Balance | Amount if Ageing days > 365 | Invoices > 1 year old |

---

## Aging Bracket Classification

Based on **Overdue days** (As on Date - Document Due Date), each invoice is categorized:

| Condition | Category | Label |
|---|---|---|
| Ar Balance < 0 | Credit/Return | "On account" |
| Overdue days < 0 | Not Yet Due | "Not Due" |
| 0 ≤ Overdue days ≤ 30 | Early Overdue | "Aging 1 to 30" |
| 30 < Overdue days ≤ 60 | 1-2 Month Overdue | "Aging 31 to 60" |
| 60 < Overdue days ≤ 90 | 2-3 Month Overdue | "Aging 61 to 90" |
| 90 < Overdue days ≤ 120 | 3-4 Month Overdue | "Aging 91 to 120" |
| 120 < Overdue days ≤ 150 | 4-5 Month Overdue | "Aging 121 to 150" |
| Overdue days > 150 | >5 Month Overdue | "Aging >=151" |

---

## Output Columns - By_Customer Sheet

### Customer Identifier Columns
| Column | Source | Notes |
|---|---|---|
| **Cust Code** | Input | Customer code |
| **Cust Name** | Input | Customer name |
| **Main Ac** | Input | Main account for grouping |
| **Cust Region** | Input | Original region field |
| **Region** | Derived | Standardized/classified region |
| **Updated Status** | Derived | Customer payment status |

### Aggregated Aging Analysis
These columns sum amounts by aging bracket for the entire customer:

| Column | Calculation |
|---|---|
| **On account** | Sum of all on-account/credit amounts |
| **Not Due** | Sum of amounts not yet due |
| **Ar Balance** | Total outstanding balance |
| **Overdue** | Sum of all aging bracket amounts (auto-calculated) |
| **Aging 1 to 30** | Sum of 1-30 days overdue amounts |
| **Aging 31 to 60** | Sum of 31-60 days overdue amounts |
| **Aging 61 to 90** | Sum of 61-90 days overdue amounts |
| **Aging 91 to 120** | Sum of 91-120 days overdue amounts |
| **Aging 121 to 150** | Sum of 121-150 days overdue amounts |
| **Aging >=151** | Sum of 151+ days overdue amounts |
| **Ageing > 365** | Sum of 365+ days old amounts |

### Quarter-Based Period Columns

#### Current Quarter (Based on Selection)
- **Q1-2026**: **01/01/2026 to 15/03/2026** (Main period)
- **Q1 tail**: **16/03/2026 to 31/03/2026** (Tail period)
- **Q1-2026 - pivot**: Read-only summary reference
- **% for Q1**: Manual entry field (enter collection % expected)
- **Actual Q1**: Formula = Q1-2026 × % for Q1 (auto-calculated)
- **Remaining % from q1**: Formula = 1 - % for Q1 (auto-calculated)
- **To add to Q2**: Formula = Remaining % × Q1-2026 (auto-calculated)

#### Future Quarters (Auto-generated based on selection)
- **Q2-2026**: **16/03/2026 to 15/06/2026** (if Q2 is future)
- **Q2 tail**: **16/06/2026 to 30/06/2026**
- **Q3-2026**: **16/06/2026 to 15/09/2026**
- **Q3 tail**: **16/09/2026 to 30/09/2026**
- **Q4-2026**: **16/09/2026 to 15/12/2026**
- **Q4 tail**: **16/12/2026 to 31/12/2026**

#### Year Columns
- **2027**: **16/12/2026 to 31/12/2027**
- **2028**: **01/01/2028 to 31/12/2028**
- **2029**: **01/01/2029 to 31/12/2029**
- **2030**: **01/01/2030 to 31/12/2030**

### Period Allocation Rules

Invoices are **allocated to quarters based on their Document Due Date**:
- Only **positive invoices** (non-credit) are allocated
- Invoice is placed in the quarter/period if: **Start Date ≤ Due Date ≤ End Date**
- **Blocked amounts** are set to 0 (not allocated to any quarter):
  - Customers with names: MINDWARE, AKLANIAT, IFIX
  - Main accounts: 12302, 12304, 12306
  - Status not in: GOOD, REGULAR, SUBSTANDARD

### Forecasting Columns

| Column | Formula | What It Means |
|---|---|---|
| **Forecast Q2** | Q2-2026 + (To add to Q2) | Predicted collection in next quarter |

**Other formula-driven columns (Remaining %, To add, Forecast)** are automatically calculated based on your % entry.

---

## Key System Rules & Conditions

### Rule 1: Date Parsing
- **As on Date** must be a valid date in cell B14
- **Document Date** and **Document Due Date** are parsed from input
- If dates are invalid or missing:
  - Missing Document Date → Assumed = As on Date
  - Missing Document Due Date → Assumed = As on Date
  - Invalid format → Error message with row numbers

### Rule 2: Region Classification
- If "Cust Region" column exists → Use classification rules
- If "Cust Code" exists → Use code-based rules
- If both exist → Combined logic
- If neither → Default to "KSA"

### Rule 3: Customer Status
- If "Customer Status" missing → Default "SUBSTANDARD"
- Empty values → Filled with "SUBSTANDARD"
- Used to **block allocation** if status ≠ (GOOD, REGULAR, SUBSTANDARD)

### Rule 4: Quarter Selection Impact
When you select a starting quarter, it affects:
1. **Which quarters appear** in output (current onward)
2. **Date ranges** for period allocations
3. **Column labels** (dynamic based on remaining quarters)

#### Q1 Selection
- Current: Q1-2026 (01/01 - 15/03)
- Tail: Q1 tail (16/03 - 31/03)
- Future: Q2, Q3, Q4 of 2026
- Beyond: 2027, 2028, 2029, 2030

#### Q2 Selection
- Current: Q2-2026 (04/01 - 15/06)
- Tail: Q2 tail (16/06 - 30/06)
- Future: Q3, Q4 of 2026
- Beyond: 2027, 2028, 2029, 2030

#### Q3 Selection
- Current: Q3-2026 (07/01 - 15/09)
- Tail: Q3 tail (16/09 - 30/09)
- Future: Q4 of 2026
- Beyond: 2027, 2028, 2029, 2030

#### Q4 Selection
- Current: Q4-2026 (10/01 - 15/12)
- Tail: Q4 tail (16/12 - 31/12)
- Future: None (moves to 2027+)
- Beyond: 2027, 2028, 2029, 2030

### Rule 5: Allocation Blocking
Amounts are **blocked from any quarter allocation** (set to 0) if ANY of these is true:
1. Customer Name contains (case-insensitive): **MINDWARE**, **AKLANIAT**, **IFIX**
2. Main Account is: **12302**, **12304**, or **12306**
3. Updated Status is NOT in: **GOOD**, **REGULAR**, **SUBSTANDARD**

Blocked amounts still appear in summary aging columns but don't get forecast into quarters.

---

## Excel Formulas in By_Customer Sheet

### Auto-Calculated Fields

#### Actual Q1 (Example for Q1)
```
=IFERROR(Q1-2026 × % for Q1, 0)
= IF(Special Account Zero, 0, IFERROR(...))
```
**What it does**: Multiplies the Q1 period balance by your entered percentage to get expected actual collection.

#### Remaining % from q1 (Example)
```
=IFERROR(1 - % for Q1, 0)
= IF(Special Account Zero, 0, IFERROR(...))
```
**What it does**: Calculates uncollected portion as percentage.

#### To add to Q2 (Example)
```
=IFERROR(Remaining % × Q1-2026, 0)
= IF(Special Account Zero, 0, IFERROR(...))
```
**What it does**: Calculates amount to add to next quarter based on uncollected amount.

#### Forecast Q2 (Example)
```
=IFERROR(Q2-2026 + To add to Q2, 0)
= IF(Special Account Zero, 0, IFERROR(...))
```
**What it does**: Sums the native Q2 amount with the carryover to predict Q2 collection.

### Special Account Zero Guard
If **Main Ac** is 12302, 12304, or 12306, all formulas return **0** regardless of amounts. This disables forecasting for controlled accounts.

---

## Example Workflow

### Your Input Excel
```
Cell B14 (As on Date): 2026-04-12
Row 16: Headers
Row 17+: Invoice data
  - INV-001: Due 2026-04-30, Amount 100,000
  - INV-002: Due 2026-05-15, Amount 50,000
  - INV-003: Due 2026-07-10, Amount 75,000
```

### You Select: Q2 as starting quarter

### Output (By_Customer):
```
Customer: ABC Company
Main Ac: 11001
Region: KSA
Status: GOOD

Q2-2026 (Apr 1 - Jun 15): 100,000 (INV-001) + 50,000 (INV-002) = 150,000
Q2 tail (Jun 16-30): 0
Q3-2026 (Jun 16 - Sep 15): 75,000 (INV-003)

% for Q2: 80% (you enter)
Actual Q2: 150,000 × 80% = 120,000 (auto formula)
Remaining % from q2: 1 - 80% = 20%
To add to Q3: 20% × 150,000 = 30,000
Forecast Q3: 75,000 + 30,000 = 105,000
```

---

## Performance Metrics

The tool displays after processing:
- **Processing Time**: How long reading and calculating took
- **Export Time**: How long generating the Excel file took
- **Total Runtime**: Total start-to-finish time

Typical performance:
- 10,000 invoices: ~2-5 seconds
- 50,000 invoices: ~5-15 seconds
- 100,000+ invoices: ~15-30 seconds

---

## Troubleshooting

### Error: "Cell B14 must contain 'As on Date'"
- **Fix**: Ensure cell B14 has a date value
- Check: No merged cells, no formatting issues blocking the read

### Error: "As on Date in B14 is not a valid date"
- **Fix**: Verify the date is in standard format (e.g., 2026-04-12 or 12/04/2026)
- **Fix**: No text or formulas, just a date value

### Error: "Missing column 'Not Due Amount'"
- **Fix**: Upload file must use the standard AR Backlog template
- Check: Your Excel has the exact columns expected

### Warning in processing (red) about invalid datetime values
- **Fix**: Check those specific row numbers listed
- **Fix**: Correct date formatting in Document Date or Document Due Date columns
- The tool will still process, but those dates are treated as missing

### Forecast formulas showing errors (#VALUE!, #NAME!)
- **Fix**: Ensure you entered numbers in the "% for Q1" (or respective quarter) column
- **Fix**: Values should be decimals (0.8 for 80%) or percentages (80%)

### All forecasts showing 0 when they shouldn't
- **Check**: Is the Main Account in the blocked list (12302, 12304, 12306)?
- **Check**: Is the Customer Status something other than GOOD/REGULAR/SUBSTANDARD?
- **Check**: Is the customer name MINDWARE, AKLANIAT, or IFIX?

---

## Summary

The **Orion Tool** is a specialized AR forecasting system that:
1. **Ingests** detailed AR backlog with invoice-level data
2. **Enriches** with aging analysis and regional classification
3. **Aggregates** by customer/account with period allocations
4. **Forecasts** collections by quarter with user-enterable percentage assumptions
5. **Outputs** three ready-to-use Excel sheets for analysis and planning

Use it to transform raw AR data into actionable customer-level forecasts with aging visibility.
