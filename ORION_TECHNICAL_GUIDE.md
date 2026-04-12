# Orion Tool - Technical Documentation
## Detailed Conditions, Calculations & Algorithm Reference

---

## Table of Contents
1. [Data Flow Architecture](#data-flow-architecture)
2. [Input Processing](#input-processing)
3. [Column-by-Column Calculation Conditions](#column-by-column-calculation-conditions)
4. [Period Allocation Logic](#period-allocation-logic)
5. [Quarter Selection System](#quarter-selection-system)
6. [Export & Formula Generation](#export--formula-generation)
7. [Edge Cases & Validation](#edge-cases--validation)

---

## Data Flow Architecture

```
Input Excel File (XLSX)
    ↓
[1] Read Excel File (openpyxl)
    - Extract first sheet name
    - Read header block (rows 1-14) to locate "As on Date"
    - Parse data from row 17 onward
    ↓
[2] Process AR File (process_ar_file)
    - Sanitize column names (decode HTML, remove NBSP)
    - Parse all dates
    - Calculate aging metrics
    - Classify regions
    - Categorize by aging bracket
    - Add 16 derived columns
    ↓ df_main (AR_Backlog sheet)
    ↓
[3] Customer Summary (customer_summary)
    - Group by Cust Code + Main Ac
    - Allocate amounts to quarters by due date
    - Apply blocking rules
    - Add forecast formulas
    ↓ df_customer (By_Customer sheet)
    ↓
[4] Invoice Summary (invoice_summary)
    - Flatten to invoice level
    - Filter and rename columns
    ↓ df_invoice (Invoice sheet)
    ↓
[5] Excel Export (fast_excel_download_multiple_with_formulas)
    - Sheet 1: AR_Backlog (all detail columns)
    - Sheet 2: By_Customer (with embedded XLSXWriter formulas)
    - Sheet 3: Invoice (formatted detail list)
    ↓
Output: processed_AR_backlog.xlsx
```

---

## Input Processing

### Phase 1: File Reading & Validation

**Function**: `process_ar_file(file)`  
**Location**: `orion/processor.py:43-165`

#### 1.1 Reading Metadata (As on Date)

```python
header_block = pd.read_excel(excel, sheet_name=sheet, header=None, nrows=14)
as_on_date = header_block.iloc[13, 1]  # Row 14 (0-indexed: 13), Column B (0-indexed: 1)
```

**Condition Chain**:
```
if pd.isna(as_on_date):
    raise ValueError("Cell B14 must contain 'As on Date'")  # Halt immediately

as_on_date = pd.to_datetime(str(as_on_date).strip(), errors="coerce")

if pd.isna(as_on_date):
    raise ValueError("As on Date in B14 is not a valid date")  # Halt immediately
```

**Valid Formats**:
- ISO: `2026-04-12`
- Excel Date serial: `46,400` (gets converted)
- Text dates: `04/12/2026`, `2026-04-12 00:00:00`

**Invalid Cases** (triggers error):
- Empty cell
- Text like "TBD", "pending"
- Corrupted date value

**Stored as**: `df.attrs["as_on_date"]` (metadata)

#### 1.2 Reading Data Rows

```python
df = pd.read_excel(excel, sheet_name=sheet, header=15, dtype=str)
# header=15 means Row 16 (0-indexed) is the header, data starts Row 17
```

**Why `dtype=str`**:
- Preserves all input as-is (dates, numbers as text)
- Allows custom parsing with error reporting
- Numeric fields are manually converted later

#### 1.3 Column Name Sanitization

```python
def sanitize_colnames(df: pd.DataFrame) -> pd.DataFrame:
    def _clean(name: str) -> str:
        s = html.unescape(str(name))  # &nbsp; → space
        s = s.replace("\u00A0", " ").strip()  # U+00A0 non-breaking space
        return s
    return df.rename(columns={c: _clean(c) for c in df.columns})
```

**Handles**:
- HTML entities: `&nbsp;` → regular space
- Unicode spaces: `\u00A0` → regular space
- Leading/trailing whitespace

### Phase 2: Date Parsing with Error Reporting

**Function**: `safe_to_datetime(series, column_name="")`  
**Location**: `orion/processor.py:29-40`

```python
def safe_to_datetime(series, column_name: str = "") -> pd.Series:
    if series is None:
        return pd.to_datetime(pd.Series([], dtype="object"), errors="coerce")
    
    # Step 1: Convert to string
    series_str = series.astype(str).str.strip()
    
    # Step 2: Normalize empty/null markers
    series_str = series_str.replace({"nan": np.nan, "": np.nan})
    
    # Step 3: Parse dates
    dt = pd.to_datetime(series_str, errors="coerce")
    
    # Step 4: Find bad rows (non-parseable but non-empty)
    bad = dt.isna() & series_str.notna()
    
    # Step 5: Report bad rows to user
    if bad.any():
        st.error(f"⚠️ Invalid datetime values in column '{column_name}':")
        for idx, val in series_str[bad].items():
            st.write(f"Row {idx + 16} → {val}")  # +16 because data starts at row 17
    
    return dt
```

**Calling Pattern**:
```python
doc_date = safe_to_datetime(df.get("Document Date"), "Document Date")
due_date = safe_to_datetime(df.get("Document Due Date"), "Document Due Date")
```

**Parsed fields**:
- Document Date (optional in input)
- Document Due Date (optional in input)

**Size Check** (handles missing columns):
```python
if len(doc_date) != len(df):
    doc_date = pd.to_datetime(pd.Series([pd.NaT] * len(df)), errors="coerce")
```
If column doesn't exist, entire series is NaT.

#### 2.1 Date Filling (Missing Value Treatment)

```python
doc_date_filled = doc_date.fillna(as_on_date)  # Missing doc date → use As on Date
due_date_filled = due_date.fillna(as_on_date)  # Missing due date → use As on Date
```

**Why**: Allows aging calculations even with missing dates.

---

## Column-by-Column Calculation Conditions

### Derived Columns Added in process_ar_file

#### Column: Ageing (Days)

**Source Code**:
```python
ar_balance = pd.to_numeric(df.get("Ar Balance"), errors="coerce").fillna(0)
doc_date = safe_to_datetime(df.get("Document Date"), "Document Date")
doc_date_filled = doc_date.fillna(as_on_date)
ageing_days = (as_on_date - doc_date_filled).dt.days
df["Ageing (Days)"] = ageing_days
```

**Formula**:
```
Ageing (Days) = (As on Date - Document Date) in days
```

**Conditions**:
- If Document Date is missing → Use As on Date → Result = 0 days
- If Document Date is in future → Result is negative
- If Document Date is valid past date → Result is positive integer

**Range**: -∞ to +∞ (typically 0-1000+ for old invoices)

**Example**:
- As on Date: 2026-04-12
- Document Date: 2026-03-15
- Result: (2026-04-12 - 2026-03-15) = 28 days

---

#### Column: Overdue days (Days)

**Source Code**:
```python
due_date = safe_to_datetime(df.get("Document Due Date"), "Document Due Date")
due_date_filled = due_date.fillna(as_on_date)
overdue_days = (as_on_date - due_date_filled).dt.days
overdue_num = pd.to_numeric(overdue_days, errors="coerce").fillna(0)
df["Overdue days (Days)"] = overdue_num
```

**Formula**:
```
Overdue days (Days) = (As on Date - Document Due Date) in days
```

**Conditions**:
- If Document Due Date is missing → Use As on Date → Result = 0 days
- If Due Date is today → Result = 0
- If Due Date is in past → Result is positive (overdue)
- If Due Date is in future → Result is negative (not yet due)

**Interpretation**:
- **Negative**: Invoice not yet due (days until due date)
- **Zero**: Due today
- **Positive**: Days overdue

**Range**: -∞ to +∞ (typically -30 to 1000+)

**Example**:
- As on Date: 2026-04-12
- Document Due Date: 2026-04-05
- Result: (2026-04-12 - 2026-04-05) = 7 days (overdue by 7 days)

---

#### Column: Region (Derived)

**Source Code**:
```python
if "Cust Region" in df.columns and "Cust Code" in df.columns:
    region_series = classify_region(df["Cust Region"], df["Cust Code"])
elif "Cust Region" in df.columns:
    region_series = classify_region(df["Cust Region"])
elif "Cust Code" in df.columns:
    empty_region = pd.Series([""] * len(df), index=df.index)
    region_series = classify_region(empty_region, df["Cust Code"])
else:
    region_series = pd.Series(["KSA"] * len(df), index=df.index)

df["Region (Derived)"] = region_series
```

**Logic Precedence**:
1. **Both Cust Region & Cust Code exist**: Use combined classification
2. **Only Cust Region exists**: Classify by region
3. **Only Cust Code exists**: Classify by code
4. **Neither exists**: Default to "KSA"

**Calling Function**: `classify_region(region_series, cust_code_series=None)`  
**Location**: `common/region_maps.py`

**Note**: Actual classification logic depends on your region_maps.py implementation. Defaults to KSA if unmapped.

---

#### Column: Aging Bracket (Label)

**Source Code**:
```python
conditions = [
    ar_balance < 0,                                  # [0] Negative balance
    overdue_num < 0,                                 # [1] Not yet due
    (overdue_num >= 0) & (overdue_num <= 30),       # [2] 0-30 days
    (overdue_num > 30) & (overdue_num <= 60),       # [3] 31-60 days
    (overdue_num > 60) & (overdue_num <= 90),       # [4] 61-90 days
    (overdue_num > 90) & (overdue_num <= 120),      # [5] 91-120 days
    (overdue_num > 120) & (overdue_num <= 150),     # [6] 121-150 days
    overdue_num > 150,                               # [7] >150 days
]

choices = [
    "On account",           # [0]
    "Not Due",              # [1]
    "Aging 1 to 30",        # [2]
    "Aging 31 to 60",       # [3]
    "Aging 61 to 90",       # [4]
    "Aging 91 to 120",      # [5]
    "Aging 121 to 150",     # [6]
    "Aging >=151",          # [7]
]

aging_bracket_label = np.select(conditions, choices, default="")
df["Aging Bracket (Label)"] = aging_bracket_label
```

**Classification Logic** (FIRST MATCH WINS):

| Priority | Condition | Result | Notes |
|---|---|---|---|
| 1 | `Ar Balance < 0` | "On account" | Credits/returns (negative amounts) |
| 2 | `Overdue days < 0` | "Not Due" | Invoice not yet due |
| 3 | `0 ≤ Overdue days ≤ 30` | "Aging 1 to 30" | 0-30 days past due |
| 4 | `30 < Overdue days ≤ 60` | "Aging 31 to 60" | 31-60 days past due |
| 5 | `60 < Overdue days ≤ 90` | "Aging 61 to 90" | 61-90 days past due |
| 6 | `90 < Overdue days ≤ 120` | "Aging 91 to 120" | 91-120 days past due |
| 7 | `120 < Overdue days ≤ 150` | "Aging 121 to 150" | 121-150 days past due |
| 8 | `Overdue days > 150` | "Aging >=151" | >150 days past due |
| — | No match | "" | (empty string) |

**Boundary Notes**:
- Exactly 30 days → "Aging 1 to 30"
- Exactly 31 days → "Aging 31 to 60"
- Negative overdue days → "Not Due" (processed before aging brackets)

---

#### Column: Updated Status

**Source Code**:
```python
if "Customer Status" in df.columns:
    updated_status = df["Customer Status"].fillna("").replace("", "SUBSTANDARD")
else:
    updated_status = pd.Series(["SUBSTANDARD"] * len(df), index=df.index)

df["Updated Status"] = updated_status
```

**Logic**:
1. If "Customer Status" column exists:
   - Extract values
   - Replace NaN with empty string
   - Replace empty string with "SUBSTANDARD"
   - Keep non-empty values as-is
2. If "Customer Status" column missing:
   - Fill entire series with "SUBSTANDARD"

**Result Values** (typical):
- "GOOD"
- "REGULAR"
- "SUBSTANDARD" (default)
- Any other value from input (preserved)

---

#### Column: Invoice Value (Derived)

**Source Code**:
```python
invoice_value = ar_balance.clip(lower=0)
df["Invoice Value (Derived)"] = invoice_value
```

**Formula**:
```
Invoice Value = MAX(Ar Balance, 0)
```

**Logic**:
- If Ar Balance ≥ 0 → Use Ar Balance
- If Ar Balance < 0 → Use 0 (exclude credits)

**Purpose**: Positive-only amounts for aging allocations.

---

#### Column: On Account (Derived)

**Source Code**:
```python
on_account_amount = ar_balance.clip(upper=0)
df["On Account (Derived)"] = on_account_amount
```

**Formula**:
```
On Account = MIN(Ar Balance, 0)  [as absolute value context]
```

**Logic**:
- If Ar Balance ≤ 0 → Use Ar Balance (negative value)
- If Ar Balance > 0 → Use 0

**Purpose**: Track credits/returns/on-account adjustments.

---

#### Column: Not Due (Derived)

**Source Code**:
```python
not_due_amount = np.where(overdue_num > 0, 0, invoice_value)
df["Not Due (Derived)"] = not_due_amount
```

**Formula**:
```
Not Due Amount = IF(Overdue days > 0, 0, Invoice Value)
```

**Logic**:
- If Overdue days > 0 (past due) → 0
- If Overdue days ≤ 0 (not yet due) → Invoice Value
- Uses Invoice Value (clipped to ≥0)

**Conditions**:
- Overdue days = 0 (due today) → Amount = Invoice Value (not yet past due)
- Overdue days < 0 (future due date) → Amount = Invoice Value
- Overdue days > 0 (already past due) → Amount = 0

---

#### Columns: Aging 1 to 30, 31 to 60, 61 to 90, 91 to 120, 121 to 150, >=151

**Source Code** (Complex allocation logic):
```python
BP, BK = invoice_value, overdue_num  # Alias for clarity

# Allocate full amount to its aging bracket
amt_ge151 = np.where(BK > 150, BP, 0)                    # >150 days
amt_121_150 = np.where(BK > 120, BP, 0) - amt_ge151      # 121-150 days
amt_91_120 = np.where(BK > 90, BP, 0) - amt_121_150 - amt_ge151      # 91-120 days
amt_61_90 = np.where(BK > 60, BP, 0) - amt_91_120 - amt_121_150 - amt_ge151  # 61-90 days
amt_31_60 = np.where(BK > 30, BP, 0) - amt_61_90 - amt_91_120 - amt_121_150 - amt_ge151  # 31-60 days
amt_1_30 = np.where(BK >= 0, BP, 0) - amt_31_60 - amt_61_90 - amt_91_120 - amt_121_150 - amt_ge151  # 0-30 days

# Clip all to ≥ 0 (safety: ensure no negatives from subtraction)
for arr in [amt_1_30, amt_31_60, amt_61_90, amt_91_120, amt_121_150, amt_ge151]:
    np.maximum(arr, 0, out=arr)

df["Aging 1 to 30 (Amount)"] = amt_1_30
df["Aging 31 to 60 (Amount)"] = amt_31_60
df["Aging 61 to 90 (Amount)"] = amt_61_90
df["Aging 91 to 120 (Amount)"] = amt_91_120
df["Aging 121 to 150 (Amount)"] = amt_121_150
df["Aging >=151 (Amount)"] = amt_ge151
```

**Algorithm Logic** (Cumulative Thresholding):
For each invoice:
1. **If Overdue > 150**: Full amount goes to "Aging >=151", 0 to all others
2. **If 120 < Overdue ≤ 150**: Full amount goes to "Aging 121 to 150"
3. **If 90 < Overdue ≤ 120**: Full amount goes to "Aging 91 to 120"
4. **If 60 < Overdue ≤ 90**: Full amount goes to "Aging 61 to 90"
5. **If 30 < Overdue ≤ 60**: Full amount goes to "Aging 31 to 60"
6. **If 0 ≤ Overdue ≤ 30**: Full amount goes to "Aging 1 to 30"
7. **If Overdue < 0**: 0 to all (handled by Not Due)
8. **If Overdue = 0**: Goes to "Aging 1 to 30"

**Key Property**: Each invoice appears in exactly ONE aging bracket (no splitting).

**Boundary Examples**:
- Overdue = 30 days → "Aging 1 to 30" (inclusive of 30)
- Overdue = 31 days → "Aging 31 to 60"
- Overdue = 150 days → "Aging 121 to 150"
- Overdue = 151 days → "Aging >=151"

---

#### Column: Ageing > 365 (Amt)

**Source Code**:
```python
df["Ageing > 365 (Amt)"] = np.where(df["Ageing (Days)"] > 365, ar_balance, 0)
```

**Formula**:
```
Ageing > 365 = IF(Ageing (Days) > 365, Ar Balance, 0)
```

**Logic**:
- If Ageing Days > 365 (older than 1 year) → Use original Ar Balance (including credits)
- Otherwise → 0

**Note**: Uses original `ar_balance`, not clipped `invoice_value` (preserves negatives).

---

#### Column: Ar Balance (Copy)

**Source Code**:
```python
df["Ar Balance (Copy)"] = ar_balance
```

**Purpose**: Preserved copy of original Ar Balance for downstream calculations (especially in customer summary).

---

### End of process_ar_file

**Output Storage**:
```python
df.attrs["as_on_date"] = as_on_date  # Metadata for downstream
return df  # 16 derived columns added to original columns
```

---

## Period Allocation Logic

### Function: customer_summary

**Location**: `orion/processor.py:168-369`

#### Phase 1: Setup & Column Preparation

```python
def customer_summary(df, selected_quarter="Q1"):
    df = sanitize_colnames(df)
    out = df.copy()
    out = out.loc[:, ~out.columns.duplicated(keep="last")]  # Remove duplicate columns
    cfg = build_customer_output_config(selected_quarter)  # Get period definitions
```

**Config returned by build_customer_output_config**:
- `selected_quarter`: "Q1", "Q2", "Q3", or "Q4"
- `active_quarters`: List of remaining quarters (e.g., if Q2 selected: ["Q2", "Q3", "Q4"])
- `current_period_label`: "Q2-2026"
- `percent_label`: "% for Q2"
- `actual_label`: "Actual Q2"
- `remaining_label`: "Remaining % from q2"
- `to_add_label`: "To add to Q3"
- `forecast_label`: "Forecast Q3"
- `later_quarter_labels`: ["Q3-2026", "Q4-2026"]
- `year_labels`: ["2027", "2028", "2029", "2030"]
- `tail_labels`: [tail date ranges for active quarters]

#### Phase 2: Field Standardization

```python
out["Cust Code"] = out.get("Cust Code", "").astype(str).str.strip()
out["Main Ac"] = out.get("Main Ac", "").fillna("").astype(str).str.strip()
out["Cust Name"] = out.get("Cust Name", "").fillna("").astype(str).str.strip()
```

**What**: Convert to string, strip whitespace, fill NaN with empty.

#### Phase 3: Region & Status Resolution

```python
if "Region" not in out.columns:
    if "Region (Derived)" in out.columns:
        out["Region"] = out["Region (Derived)"]
    elif "Cust Region" in out.columns:
        out["Region"] = out["Cust Region"]
    else:
        out["Region"] = ""
        
if "Customer Status" omitted:
    (Already handled in process_ar_file as "Updated Status")
```

#### Phase 4: Numeric Field Preparation

```python
for c in [
    "On Account (Derived)",
    "Aging 1 to 30 (Amount)",
    "Aging 31 to 60 (Amount)",
    "Aging 61 to 90 (Amount)",
    "Aging 91 to 120 (Amount)",
    "Aging 121 to 150 (Amount)",
    "Aging >=151 (Amount)",
    "Ageing > 365 (Amt)",
    "Overdue days (Days)",
    "Not Due Amount",
    "Ar Balance (Copy)",
]:
    out[c] = pd.to_numeric(out.get(c, 0), errors="coerce").fillna(0)
```

**Effect**: Ensures numeric columns are numeric, defaults to 0 if missing.

#### Phase 5: Due Date Parsing for Period Allocation

```python
if "Document Due Date" in out.columns:
    due_raw = out["Document Due Date"].astype(str).str.strip()
    due_raw = due_raw.replace("\u00A0", " ", regex=False)  # Non-breaking space
    due_raw = due_raw.str.replace(r"[^\x00-\x7F]", "", regex=True)  # Remove non-ASCII
    due_raw = due_raw.str.replace(r"\s+\d{2}:\d{2}:\d{2}$", "", regex=True)  # Remove time
    due_dt = pd.to_datetime(due_raw, errors="coerce")
else:
    due_dt = pd.Series([pd.NaT] * len(out))
```

**Why the cleaning**:
- Unicode spaces interfere with date parsing
- Time components (e.g., `2026-04-12 14:30:00`) are removed
- Non-ASCII characters are stripped
- Parse result is a datetime series with NaT for unparseable values

---

#### Phase 6: Period Mapping (Based on Selected Quarter)

**Source Code**:
```python
yr = 2026
period_map = {
    "Q1": [
        ("Q1-2026", (pd.Timestamp(yr, 1, 1), pd.Timestamp(yr, 3, 15))),
        ("Q1 tail", (pd.Timestamp(yr, 3, 16), pd.Timestamp(yr, 3, 31))),
        ("Q2-2026", (pd.Timestamp(yr, 3, 16), pd.Timestamp(yr, 6, 15))),
        ("Q2 tail", (pd.Timestamp(yr, 6, 16), pd.Timestamp(yr, 6, 30))),
        ("Q3-2026", (pd.Timestamp(yr, 6, 16), pd.Timestamp(yr, 9, 15))),
        ("Q3 tail", (pd.Timestamp(yr, 9, 16), pd.Timestamp(yr, 9, 30))),
        ("Q4-2026", (pd.Timestamp(yr, 9, 16), pd.Timestamp(yr, 12, 15))),
        ("Q4 tail", (pd.Timestamp(yr, 12, 16), pd.Timestamp(yr, 12, 31))),
        ("2027", (pd.Timestamp(2027, 12, 16), pd.Timestamp(2027, 12, 31))),
        ("2028", (pd.Timestamp(2028, 1, 1), pd.Timestamp(2028, 12, 31))),
        ("2029", (pd.Timestamp(2029, 1, 1), pd.Timestamp(2029, 12, 31))),
        ("2030", (pd.Timestamp(2030, 1, 1), pd.Timestamp(2030, 12, 31))),
    ],
    "Q2": [
        ("Q2-2026", (pd.Timestamp(yr, 4, 1), pd.Timestamp(yr, 6, 15))),
        ("Q2 tail", (pd.Timestamp(yr, 6, 16), pd.Timestamp(yr, 6, 30))),
        # ... Q3, Q4, years
    ],
    "Q3": [
        ("Q3-2026", (pd.Timestamp(yr, 7, 1), pd.Timestamp(yr, 9, 15))),
        ("Q3 tail", (pd.Timestamp(yr, 9, 16), pd.Timestamp(yr, 9, 30))),
        # ... Q4, years
    ],
    "Q4": [
        ("Q4-2026", (pd.Timestamp(yr, 10, 1), pd.Timestamp(yr, 12, 15))),
        ("Q4 tail", (pd.Timestamp(yr, 12, 16), pd.Timestamp(yr, 12, 31))),
        # ... years only
    ],
}
```

**Period Definition**:
- Main period and tail for each quarter (15-day split on 15th/16th)
- Year periods (Jan 1 - Dec 31)
- 2027 tail special case (starts from Dec 16, 2026)

---

#### Phase 7: Blocking Rules (Critical for Allocation)

```python
ZERO_QUARTER_CUSTOMER_KEYWORDS = ("MINDWARE", "AKLANIAT", "IFIX")
ZERO_COLLECTION_MAIN_ACCOUNTS = {"12302", "12304", "12306"}

# Block if customer name matches (case-insensitive)
blocked_customer = out["Cust Name"].str.upper().str.contains(
    "|".join(ZERO_QUARTER_CUSTOMER_KEYWORDS), na=False
)

# Block if main account is in list
blocked_main_account = out["Main Ac"].isin(ZERO_COLLECTION_MAIN_ACCOUNTS)

# Block if status not in allowed list (case-insensitive)
allowed_statuses = ["GOOD", "REGULAR", "SUBSTANDARD"]
blocked_status = ~out["Updated Status"].str.upper().isin(allowed_statuses)

# Create mask: block if ANY condition is true
quarter_amount = out["Ar Balance (Copy)"].clip(lower=0).where(
    ~(blocked_customer | blocked_main_account | blocked_status), 
    0
)
```

**Blocking conditions** (using OR logic):
1. Customer name (uppercase) **contains** "MINDWARE" OR "AKLANIAT" OR "IFIX"
2. Main Account **equals** "12302" OR "12304" OR "12306"
3. Updated Status (uppercase) **not in** ["GOOD", "REGULAR", "SUBSTANDARD"]

**Effect**: If ANY is true, `quarter_amount = 0` for that row (no allocation to quarters).

**Note**: Blocked amounts still appear in aging summaries, just don't get forecast into quarters.

---

#### Phase 8: Period Allocation

```python
for col, (start, end) in period_map[selected_quarter]:
    out[col] = np.where((due_dt >= start) & (due_dt <= end), quarter_amount, 0)
```

**Allocation Rule** (for each period in period_map):
```
Column Value = IF(Document Due Date >= Period Start AND Document Due Date <= Period End, quarter_amount, 0)
```

**Boundary Conditions**:
- Due Date = Period Start → Included (>=)
- Due Date = Period End → Included (<=)
- Due Date outside range → 0

**Examples**:
- Due Date: 2026-04-12, Period: Q2-2026 (Apr 1 - Jun 15)
  - Check: 2026-04-12 >= 2026-04-01 AND 2026-04-12 <= 2026-06-15 → TRUE
  - Result: `quarter_amount` (if not blocked)

- Due Date: 2026-06-16, Period: Q2-2026 (Apr 1 - Jun 15)
  - Check: 2026-06-16 >= 2026-04-01 AND 2026-06-16 <= 2026-06-15 → FALSE
  - Result: 0 (goes to Q2 tail instead if it matches)

---

#### Phase 9: Aggregation by Customer

```python
agg_map = {
    "Cust Name": ("Cust Name", "first"),
    "Cust Region": ("Cust Region", "first"),
    "Region": ("Region", "first"),
    "Updated Status": ("Updated Status", "first"),
    "On Account (Derived)": ("On Account (Derived)", "sum"),
    "Not Due Amount": ("Not Due Amount", "sum"),
    "AR Balance": ("Ar Balance (Copy)", "sum"),
    "Overdue days (Days)": ("Overdue days (Days)", "sum"),
    "Aging 1 to 30 (Amount)": ("Aging 1 to 30 (Amount)", "sum"),
    "Aging 31 to 60 (Amount)": ("Aging 31 to 60 (Amount)", "sum"),
    "Aging 61 to 90 (Amount)": ("Aging 61 to 90 (Amount)", "sum"),
    "Aging 91 to 120 (Amount)": ("Aging 91 to 120 (Amount)", "sum"),
    "Aging 121 to 150 (Amount)": ("Aging 121 to 150 (Amount)", "sum"),
    "Aging >=151 (Amount)": ("Aging >=151 (Amount)", "sum"),
    "Ageing > 365 (Amt)": ("Ageing > 365 (Amt)", "sum"),
}

# Add period columns to agg_map (all summed)
for col, _ in period_map[selected_quarter]:
    agg_map[col] = (col, "sum")

grouped = out.groupby(["Cust Code", "Main Ac"], as_index=False).agg(**agg_map)
```

**Grouping**:
- Group by (`Cust Code`, `Main Ac`) combination
- First-appear fields: Cust Name, Region, Status (use "first" because same for all rows in group)
- Summed fields: All amounts

**Result**: One row per (Cust Code, Main Ac) pair.

---

#### Phase 10: Overdue Recalculation

```python
amount_buckets = [
    "Aging 1 to 30 (Amount)",
    "Aging 31 to 60 (Amount)",
    "Aging 61 to 90 (Amount)",
    "Aging 91 to 120 (Amount)",
    "Aging 121 to 150 (Amount)",
    "Aging >=151 (Amount)",
]

present = [c for c in amount_buckets if c in grouped.columns]
grouped["Overdue days (Days)"] = grouped[present].sum(axis=1) if present else 0
```

**Logic**:
- Sum all aging brackets to get total overdue amount
- This is stored back in "Overdue days (Days)" column (misleading name, but by design)
- Overwrites the detail-level sum (now represents total overdue amount, not days)

---

#### Phase 11: Column Renaming

```python
rename_final = {
    "On Account (Derived)": "On account",
    "Not Due Amount": "Not Due",
    "AR Balance": "Ar Balance",
    "Overdue days (Days)": "Overdue",
    "Aging 1 to 30 (Amount)": "Aging 1 to 30",
    "Aging 31 to 60 (Amount)": "Aging 31 to 60",
    "Aging 61 to 90 (Amount)": "Aging 61 to 90",
    "Aging 91 to 120 (Amount)": "Aging 91 to 120",
    "Aging 121 to 150 (Amount)": "Aging 121 to 150",
    "Aging >=151 (Amount)": "Aging >=151",
    "Ageing > 365 (Amt)": "Ageing > 365",
}

grouped = grouped.rename(columns=rename_final)
```

---

#### Phase 12: Dynamic Manual Columns (User Entry)

```python
dynamic_manual_cols = [
    cfg["percent_label"],      # "% for Q2" (example)
    cfg["actual_label"],       # "Actual Q2"
    cfg["remaining_label"],    # "Remaining % from q2"
    cfg["to_add_label"],       # "To add to Q3"
    cfg["forecast_label"],     # "Forecast Q3"
]

for col in dynamic_manual_cols:
    if col not in grouped.columns:
        grouped[col] = 0
```

**Purpose**: Create columns for user manual entry and formula-fed calculations.
- Initialized to 0
- Will be populated with formulas in Excel export

---

#### Phase 13: Final Column Ordering

```python
final_order = [
    "Cust Code",
    "Cust Name",
    "Main Ac",
    "Cust Region",
    "Region",
    "Updated Status",
    "On account",
    "Not Due",
    "Ar Balance",
    "Overdue",
    "Aging 1 to 30",
    "Aging 31 to 60",
    "Aging 61 to 90",
    "Aging 91 to 120",
    "Aging 121 to 150",
    "Aging >=151",
    "Ageing > 365",
    cfg["current_pivot_label"],         # "Q2-2026 - pivot"
    cfg["current_period_label"],        # "Q2-2026"
    QUARTER_TAIL_LABELS_2026[selected_quarter],  # "16/06/2026..30/06/2026"
]

final_order.extend(dynamic_manual_cols)  # % for Q2, Actual Q2, etc.

for label in cfg["later_quarter_labels"]:  # Q3-2026, Q4-2026, ...
    final_order.append(label)
    final_order.append(QUARTER_TAIL_LABELS_2026[label.split("-")[0]])  # Tail

final_order.extend(cfg["year_labels"])  # 2027, 2028, 2029, 2030

# Ensure all columns exist (fill missing with 0)
for c in final_order:
    if c not in grouped.columns:
        grouped[c] = 0

return grouped[final_order]
```

---

## Quarter Selection System

### build_customer_output_config Function

**Location**: `common/quarter_utils.py:25-46`

```python
def build_customer_output_config(selected_quarter: str) -> dict:
    idx = QUARTER_ORDER.index(selected_quarter)  # 0-3 for Q1-Q4
    active_quarters = QUARTER_ORDER[idx:]         # Remaining quarters
    tail_labels = [QUARTER_TAIL_LABELS_2026[q] for q in active_quarters]
    next_label = next_period_label(selected_quarter)
    next_display_label = "2027" if selected_quarter == "Q4" else QUARTER_ORDER[idx + 1]

    return {
        "selected_quarter": selected_quarter,
        "active_quarters": active_quarters,
        "current_period_label": f"{selected_quarter}-2026",
        "current_pivot_label": f"{selected_quarter}-2026 - pivot",
        "percent_label": f"% for {selected_quarter}",
        "actual_label": f"Actual {selected_quarter}",
        "remaining_label": f"Remaining % from {selected_quarter.lower()}",
        "to_add_label": f"To add to {next_display_label}",
        "forecast_label": f"Forecast {next_display_label}",
        "next_period_label": next_label,
        "tail_labels": tail_labels,
        "later_quarter_labels": [f"{q}-2026" for q in active_quarters[1:]],
        "year_labels": ["2027", "2028", "2029", "2030"],
    }
```

### Selected Quarter Examples

#### Q1 Selected
```
idx = 0
active_quarters = ["Q1", "Q2", "Q3", "Q4"]
current_period_label = "Q1-2026"
percent_label = "% for Q1"
next_display_label = "Q2"
to_add_label = "To add to Q2"
forecast_label = "Forecast Q2"
later_quarter_labels = ["Q2-2026", "Q3-2026", "Q4-2026"]
```

#### Q2 Selected
```
idx = 1
active_quarters = ["Q2", "Q3", "Q4"]
current_period_label = "Q2-2026"
percent_label = "% for Q2"
next_display_label = "Q3"
to_add_label = "To add to Q3"
forecast_label = "Forecast Q3"
later_quarter_labels = ["Q3-2026", "Q4-2026"]
```

#### Q4 Selected
```
idx = 3
active_quarters = ["Q4"]
current_period_label = "Q4-2026"
percent_label = "% for Q4"
next_display_label = "2027"  # Special case
to_add_label = "To add to 2027"
forecast_label = "Forecast 2027"
later_quarter_labels = []  # No more quarters in 2026
```

---

## Export & Formula Generation

### Function: fast_excel_download_multiple_with_formulas

**Location**: `orion/export.py:73-176`

#### Sheet 1: AR_Backlog
```python
ws_main = wb.add_worksheet("AR_Backlog")
main_df = write_sheet(ws_main, df_main)
```

**Content**: All columns from process_ar_file (original + derived).
**Rows**: One per invoice.
**Formatting**: Date columns formatted as dd/mm/yyyy.

#### Sheet 2: By_Customer (with XLSXWriter Formulas)

```python
ws_cust = wb.add_worksheet("By_Customer")
cust_df = normalize_all_date_strings(df_customer.copy()).fillna("")

col_map = build_col_map(cust_df)  # Column name → Excel letter mapping

def idx(name):
    return list(cust_df.columns).index(name)  # Get 0-indexed column position
```

**Formula Writing Logic**:

```python
for r_idx, row in enumerate(cust_df.to_numpy(), start=1):
    excel_row = r_idx + 1  # Excel row number (1-indexed with header)
    ws_cust.write_row(r_idx, 0, row.tolist())  # Write data

    # Column lookups
    current_period = col_map.get(cfg["current_period_label"])
    percent = col_map.get(cfg["percent_label"])
    remaining = col_map.get(cfg["remaining_label"])
    to_add = col_map.get(cfg["to_add_label"])
    next_period = col_map.get(cfg["next_period_label"])
    main_ac = col_map.get("Main Ac")

    # Special account guard (zeros out formulas for 12302, 12304, 12306)
    collection_zero_guard = None
    if main_ac:
        guards = [f'${main_ac}{excel_row}="{value}"' for value in ZERO_COLLECTION_MAIN_ACCOUNTS]
        collection_zero_guard = f"OR({','.join(guards)})"
```

**Formula: Actual Collection**
```
If columns exist (current_period and percent):
    actual_formula = f"IFERROR(${current_period}{excel_row}*${percent}{excel_row},0)"
    
    With guard:
    actual_formula = f"IF({collection_zero_guard},0,{actual_formula})"
    = IF(OR($Main_Ac_Row="12302",$Main_Ac_Row="12304",$Main_Ac_Row="12306"),
          0,
          IFERROR($Q2-2026_Row*$%_for_Q2_Row,0))
    
    Effect:
    - If Main Ac is 12302/12304/12306 → 0
    - Else IF error → 0
    - Else → Period Amount × Percentage
```

**Formula: Remaining %**
```
If percent column exists:
    remaining_formula = f"IFERROR(1-${percent}{excel_row},0)"
    
    With guard:
    remaining_formula = f"IF({collection_zero_guard},0,{remaining_formula})"
    
    Effect:
    - If blocked account → 0
    - Else → (1 - User % Entry)
```

**Formula: To Add (Carryover)**
```
If remaining and current_period columns exist:
    to_add_formula = f"IFERROR(${remaining}{excel_row}*${current_period}{excel_row},0)"
    
    With guard:
    to_add_formula = f"IF({collection_zero_guard},0,{to_add_formula})"
    
    Effect:
    - If blocked account → 0
    - Else → Remaining % × Period Amount
```

**Formula: Forecast Next Period**
```
If to_add and next_period columns exist:
    forecast_formula = f"IFERROR(${next_period}{excel_row}+${to_add}{excel_row},0)"
    
    With guard:
    forecast_formula = f"IF({collection_zero_guard},0,{forecast_formula})"
    
    Effect:
    - If blocked account → 0
    - Else → Next Period Native + Carryover Amount
```

**Sheet Formatting**:
```python
ws_cust.freeze_panes(1, 0)  # Freeze header row
ws_cust.autofilter(0, 0, max(1, len(cust_df)), len(cust_df.columns) - 1)  # Add filter
```

#### Sheet 3: Invoice (Detail)
```python
ws_inv = wb.add_worksheet("Invoice")
inv_df = write_sheet(ws_inv, df_invoice)
```

**Content**: From `invoice_summary()` function.

---

## Edge Cases & Validation

### Missing Columns Handling

| Column | Missing Behavior | Default |
|---|---|---|
| **Ar Balance** | `fillna(0)` | 0 |
| **Document Date** | Filled with As on Date | As on Date |
| **Document Due Date** | Filled with As on Date | As on Date |
| **Cust Region** | Classify region rules applied | "KSA" if no region info |
| **Customer Status** | Filled with "SUBSTANDARD" | "SUBSTANDARD" |
| **Not Due Amount** | Column created if missing | 0 |
| **Main Ac** | Treat as empty string | "" |

### Date Parsing Edge Cases

| Scenario | Handling |
|---|---|
| **Future Date** | Aging days = negative, Overdue = negative |
| **Same-Day Date** | Ageing = 0, Overdue = 0 (classified as "Aging 1 to 30") |
| **Null/NaT** | Filled with As on Date |
| **Invalid Format** | Coerced to NaT, error reported to user, filled with As on Date |

### Numeric Conversion

All numeric conversions use `pd.to_numeric(..., errors="coerce").fillna(0)`:
- Non-numeric strings → NaN → 0
- Empty strings → NaN → 0
- Valid numbers → kept as-is
- Negatives → preserved

### Empty Dataframe Handling

```python
if df is None or df.empty:
    return df  # Return as-is (short-circuit)
```

In normalize_all_date_strings and coerce_export_dates.

---

## Summary

This document details:
1. **File reading & As on Date validation**
2. **Date parsing with user error reporting**
3. **16 derived columns** with exact conditions
4. **Region classification** with fallback rules
5. **Aging bracket allocation** (each invoice → exactly 1 bracket)
6. **Period allocation** (quarter/date-based grouping)
7. **Customer grouping** (Cust Code + Main Ac)
8. **Excel formula generation** with guarded calculations
9. **Edge case handling** throughout pipeline

All logic is deterministic and repeatable given the same input file and As on Date.
