# DAY 3

## Chapter 4: IF, IFS and SWITCH in Excel

### ❓ IF()
Tests a specified condition and returns one value if **TRUE**, another if **FALSE**.

**Syntax:**
```excel
=IF(logical_test, value_if_true, value_if_false)
```

**Example 1 — Calculate employee bonus:**
```excel
=IF(B2>10000, B2*5%, 0)
```
> If sales > $10,000 → 5% bonus; else → $0

**Example 2 — Assign grades based on test scores:**
```excel
=IF(A2>=90, "A", IF(A2>=80, "B", "C"))
```

---

### ❓ IFS()
Checks multiple conditions and returns a value for the **first TRUE** condition.

**Syntax:**
```excel
=IFS(condition1, value1, condition2, value2, ...)
```

**Example 1 — Assign letter grades:**
```excel
=IFS(A2>=90, "A", A2>=80, "B", A2>=70, "C", A2>=60, "D", A2<60, "F")
```

**Example 2 — Apply discount tiers based on price:**
```excel
=IFS(B2<10, B2, B2<=20, B2*0.95, B2>20, B2*0.90)
```

---

### 🔀 SWITCH()
Evaluates an expression against a list of cases and returns the result for the **first match**.

**Syntax:**
```excel
=SWITCH(expression, value1, result1, value2, result2, ..., [default])
```

**Example 1 — Calculate commission based on sales tiers:**
```excel
=SWITCH(TRUE, B2<5000, B2*1%, B2<=10000, B2*2%, B2>10000, B2*3%)
```

**Example 2 — Categorize countries by continent:**
```excel
=SWITCH(A2, "USA","North America", "Brazil","South America", "France","Europe", "Unknown")
```

---

## Chapter 5: Conditional Formatting and Data Validation

### 🎨 Conditional Formatting

Applies formatting to cells based on **specific conditions or criteria**.

**Steps to Apply:**
1. Select the cells to format
2. Go to **Home** tab → click **Conditional Formatting**
3. Choose the type: **Highlight Cells Rules**, **Top/Bottom Rules**, **Color Scales**, etc.
4. Set your conditions and formatting
5. Click **OK**

**Example 1 — Highlight negative numbers:**
1. Select the range → **Conditional Formatting** → **Highlight Cells Rules** → **Less Than**
2. Enter `0` in the value box → Choose a fill color → Click **OK**

**Example 2 — Create a heat map for sales data:**
1. Select the range → **Conditional Formatting** → **Color Scales**
2. Choose a color scale (e.g., Red-Yellow-Green)
3. Click **OK**

---

### ✅ Data Validation

Sets **rules or constraints** on data entered into a cell to ensure accuracy.

**Steps to Apply:**
1. Select the cell(s) for validation
2. Go to the **Data** tab → click **Data Validation**
3. Choose validation type (Whole Number, Decimal, List, Date, Custom)
4. Set the validation criteria
5. Customize input message and error alert
6. Click **OK**

**Validation Types:**

| Type | Use Case | Example |
|------|----------|---------|
| **Whole Number** | Restrict to integers only | Allow only values between 1–100 |
| **Decimal** | Allow decimal values within range | Prices between 0.01–999.99 |
| **List** | Drop-down selection | Choose from a predefined list |
| **Date** | Restrict to a date range | Only accept dates in 2024 |
| **Custom** | Custom formula validation | `=SUM(A1:A5)>100` |

---

## Chapter 6: Working with Excel Tables

### Creating an Excel Table

1. Select your data (including headers)
2. Go to **Insert** tab → click **Table**
3. Confirm the range and check **"My table has headers"**
4. Click **OK**

### Benefits of Excel Tables

| Benefit | Description |
|---------|-------------|
| **Structured Format** | Organizes data for easier analysis |
| **Easy Sorting & Filtering** | Built-in drop-down filters on every column |
| **Built-in Data Analysis** | Supports PivotTables and PivotCharts |
| **Automatic Formatting** | Predefined styles apply automatically |
| **Easy Data Entry** | Table expands automatically when new rows are added |

### Working with Tables — Example

1. Select the entire dataset including headers
2. **Insert** → **Table** → confirm range → **OK**
3. To sort by "Sales" column (descending): click the drop-down arrow on **Sales** → **Sort Largest to Smallest**
4. To filter by product: click the drop-down on **Product** → select desired products

---

## Day 3 Assignment — Logic, Formatting & Tables

### Task 1: IF, IFS and SWITCH
Using your Student Records sheet:

1. Use `=IF()` to add a **"Pass/Fail"** column — students with a score of 50 or above Pass, others Fail
2. Use `=IF()` with a nested IF to add a **"Grade"** column:
   - 80 and above = A
   - 65–79 = B
   - 50–64 = C
   - Below 50 = F
3. Rewrite the Grade formula using `=IFS()`
4. Create a new column called **"Performance"** and use `=SWITCH()` to label grades as: A = "Excellent", B = "Good", C = "Average", F = "Poor"

### Task 2: Conditional Formatting
Using your Student Records sheet:

1. Highlight all **"Fail"** scores in red
2. Highlight all **"A" grades** in green
3. Apply a **Color Scale** to the Score column to create a heat map effect
4. Use **Top/Bottom Rules** to highlight the top 5 scores

### Task 3: Data Validation
Create a new sheet called **"Registration Form"** and apply the following validation rules:

1. Age column — only allow whole numbers between **16 and 60**
2. Score column — only allow decimals between **0 and 100**
3. Department column — create a **drop-down list** with options: Sales, Marketing, Finance, HR, IT
4. Date column — only allow dates within the **current year**
5. Add a custom **error message** for each validation rule

### Task 4: Excel Tables
1. Convert your Sales Data into a proper **Excel Table**
2. Sort the table by Total Sales (largest to smallest)
3. Filter the table to show only one region
4. Add a **Total Row** to the table and calculate the sum of Total Sales
5. Apply a table style of your choice

---
