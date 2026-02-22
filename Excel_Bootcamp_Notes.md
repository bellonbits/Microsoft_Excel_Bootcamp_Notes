# 📊 Microsoft Excel Bootcamp Notes
### 4-Day Intensive Training Guide
---

> **Trainer:** Peter Gatitu | [gatituportfolio.netlify.app/](https://gatituportfolio.netlify.app/) | @gatitu_mwangi

---

## 📅 Bootcamp Schedule Overview

| Day | Topics Covered |
|-----|---------------|
| **Day 1** | Introduction to Excel + Managing Data in Cells & Ranges |
| **Day 2** | Excel Functions (SUM, AVERAGE, LOOKUP, Logical, Text) |
| **Day 3** | IF/IFS/SWITCH + Conditional Formatting + Data Validation + Tables |
| **Day 4** | Charts, PivotTables, What-If Analysis & Data Cleaning |

---

# DAY 1

## Chapter 1: What is Excel?

Microsoft Excel is a **spreadsheet software** developed by Microsoft, widely used for organizing, analyzing, and visualizing data.

### Key Terminology

| Term | Definition |
|------|-----------|
| **Workbook** | A file that contains one or more worksheets |
| **Worksheet** | A single page/sheet in a workbook where data is entered, analysed and manipulated |
| **Cell** | A single rectangular box on a worksheet where data can be entered or displayed |
| **Range** | A group of two or more adjacent cells in a worksheet |
| **Formula** | An equation that performs calculations on values in one or more cells |
| **Function** | A predefined formula that performs a specific calculation (e.g. SUM, AVERAGE) |
| **Chart** | A graphical representation of data from a worksheet |
| **Pivot Table** | A powerful tool for summarizing, analyzing, and presenting large amounts of data |
| **Conditional Formatting** | A feature to apply formatting to cells based on certain conditions or criteria |
| **Data Validation** | A feature to restrict or control the type of data entered into a cell |
| **Power Pivot** | A data modeling tool for creating relationships between tables and advanced analysis |

---

## Chapter 2: Managing Data in Cells and Ranges

### 🔃 Sorting

Sorting allows you to arrange data in a specific order based on certain criteria.

**How to Sort:**
1. Select the range of cells you want to sort
2. Click **"Sort & Filter"** on the **Home** tab
3. Select **"Sort A to Z"** (ascending) or **"Sort Z to A"** (descending)
4. For complex sorting, select **"Custom Sort"** to sort by multiple columns

**Example 1 — Sort names alphabetically:**
1. Select the range of cells containing the names
2. Click **Sort & Filter** → **Sort A to Z**

**Example 2 — Sort sales figures by month:**
1. Select the range of cells with sales figures
2. Click **Sort & Filter** → **Custom Sort**
3. Set **Month** as the first sorting column → **Sort A to Z**

---

### 🔍 Filtering

Filtering allows you to selectively display only certain data based on specific criteria.

**How to Filter:**
1. Select the range of cells to filter
2. Click **"Filter"** on the **Data** tab
3. Use the drop-down arrows in the column headers to set your filter criteria

**Example 1 — Filter orders by customer name:**
1. Select the data range → Click **Filter** on the Data tab
2. Click the drop-down arrow in the **"Customer Name"** column
3. Select the customer name to filter by

**Example 2 — Filter sales figures by date range:**
1. Select the data range → Click **Filter**
2. Click the drop-down on the **"Date"** column → Select **"Date Filters"** → **"Between"**
3. Enter start and end dates

---

### 🧮 Calculating with Formulas

Formulas always start with an **equals sign (=)**.

**Example 1 — Calculate average of a list:**
```excel
=AVERAGE(A1:A10)
```

**Example 2 — Total sales for a specific product:**
```excel
=SUMIF(B:B,"Product A",C:C)
```

---

# DAY 2

## Chapter 3: Excel Functions

### ➕ SUM()
Adds up a range of cells and returns the total.

```excel
=SUM(A1:A5)
```
```excel
=SUM(A1:A5, B1:B5)
```

---

### 📊 AVERAGE()
Calculates the arithmetic mean of a range of cells.

```excel
=AVERAGE(A1:A5)
```
```excel
=AVERAGE(A1:A5, B1:B5)
```

---

### 🔢 COUNT()
Counts the number of cells in a range that contain numbers.

```excel
=COUNT(A1:A5)
```
```excel
=COUNT(A1:A5, B1:B5)
```

---

### 🔺 MAX()
Returns the highest value in a range of cells.

```excel
=MAX(A1:A5)
```
```excel
=MAX(A1:A5, B1:B5)
```

---

### 🔻 MIN()
Returns the minimum value in a range of cells.

**Syntax:**
```excel
=MIN(range)
```

**Example 1:**
```excel
=MIN(A1:A10)
```

**Example 2:**
```excel
=MIN(B2:B20)
```

---

### 🔗 CONCATENATE()
Joins two or more text strings into a single string.

**Syntax:**
```excel
=CONCATENATE(text1, text2, ...)
```

**Example 1 — Join "Hello" and "World":**
```excel
=CONCATENATE(A1, " ", B1)
```

**Example 2 — Create full name from first and last name:**
```excel
=CONCATENATE(A1, " ", B1)
```

> 💡 **Modern alternative:** Use `&` operator: `=A1&" "&B1`

---

### ⬅️ LEFT()
Extracts a specified number of characters from the **beginning** of a text string.

**Syntax:**
```excel
=LEFT(text, num_chars)
```

**Example 1 — Extract first 5 characters from "Excel is awesome":**
```excel
=LEFT(A1, 5)
```
> Returns: `Excel`

**Example 2 — Extract username from email address:**
```excel
=LEFT(B2, FIND("@", B2)-1)
```

---

### ➡️ RIGHT()
Extracts a specified number of characters from the **end** of a text string.

**Syntax:**
```excel
=RIGHT(text, num_chars)
```

**Example 1 — Extract last 6 characters from "Excel is awesome":**
```excel
=RIGHT(A1, 6)
```
> Returns: `awesome`

**Example 2 — Extract last 4 digits from phone numbers:**
```excel
=RIGHT(C2, 4)
```

---

### 🔍 XLOOKUP()
A modern, flexible lookup function. Searches for a value and returns a corresponding value.

**Syntax:**
```excel
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

**Parameters:**

| Argument | Description |
|----------|-------------|
| `lookup_value` | The value you want to look up |
| `lookup_array` | The range where you want to search |
| `return_array` | The range from which to return a value |
| `match_mode` | 0 = Exact match (default); -1 = Exact or next smallest; 1 = First match; 2 = Last match |
| `search_mode` | 1 = Search from beginning (default); -1 = Search from end |

**Example 1 — Find the price of "Product A":**
```excel
=XLOOKUP("Product A", A2:A10, B2:B10)
```

**Example 2 — Find GDP for "USA" and convert currency:**
```excel
=XLOOKUP("USA", A2:A10, B2:B10) * 0.85
```

---

### 🔍 VLOOKUP()
Searches for a value in the **first column** of a range and returns a value from a specified column.

**Syntax:**
```excel
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

**Parameters:**

| Argument | Description |
|----------|-------------|
| `lookup_value` | The value to search for |
| `table_array` | The range to search in |
| `col_index_num` | Column number to return (1 = first column) |
| `range_lookup` | TRUE = approximate match; FALSE = exact match |

**Example 1 — Find salary for "John Doe":**
```excel
=VLOOKUP("John Doe", A2:B10, 2, FALSE)
```

**Example 2 — Find category for "Product A":**
```excel
=VLOOKUP("Product A", A2:B10, 2, FALSE)
```

> ⚠️ **Note:** XLOOKUP is the preferred modern alternative to VLOOKUP.

---

### ✅ AND()
Returns **TRUE** if ALL conditions are true; **FALSE** if any one is false.

**Syntax:**
```excel
=AND(condition1, condition2, ...)
```

**Example 1 — Students scoring above 80 in all 3 subjects:**
```excel
=AND(B2>80, C2>80, D2>80)
```

**Example 2 — Employees earning >$50,000 in Sales or Marketing:**
```excel
=AND(B2>50000, OR(C2="Sales", C2="Marketing"))
```

---

### ✅ OR()
Returns **TRUE** if ANY one condition is true; **FALSE** if all are false.

**Syntax:**
```excel
=OR(logical1, logical2, ...)
```

**Example 1 — Employees with a rating of 4 or 5:**
```excel
=OR(B2=4, B2=5)
```

**Example 2 — Products priced above $100 or below $50:**
```excel
=OR(B2>100, B2<50)
```

---

### 🔢 COUNTIF()
Counts the number of cells in a range that meet a single condition.

**Syntax:**
```excel
=COUNTIF(range, criteria)
```

**Example 1 — Count cells with value greater than 5:**
```excel
=COUNTIF(A1:A10, ">5")
```

**Example 2 — Count cells containing "apple":**
```excel
=COUNTIF(B1:B10, "apple")
```

---

### 🔢 COUNTIFS()
Counts cells in multiple ranges that meet multiple criteria.

**Syntax:**
```excel
=COUNTIFS(range1, criteria1, range2, criteria2, ...)
```

**Example 1 — Count cells between 5 and 10:**
```excel
=COUNTIFS(A1:A10, ">5", A1:A10, "<10")
```

**Example 2 — Count cells with "apple" and value > 5:**
```excel
=COUNTIFS(B1:B10, "apple", C1:C10, ">5")
```

---

### ➕ SUMIF()
Adds values in a range that meet a single condition.

**Syntax:**
```excel
=SUMIF(range, criteria, [sum_range])
```

**Example 1 — Sum values greater than 5:**
```excel
=SUMIF(A1:A10, ">5")
```

**Example 2 — Sum values where text is "apple":**
```excel
=SUMIF(B1:B10, "apple", C1:C10)
```

---

### ➕ SUMIFS()
Adds values in multiple ranges that meet multiple criteria.

**Syntax:**
```excel
=SUMIFS(sum_range, range1, criteria1, range2, criteria2, ...)
```

**Example 1 — Sum values between 5 and 10:**
```excel
=SUMIFS(A1:A10, A1:A10, ">5", A1:A10, "<10")
```

**Example 2 — Sum values where text is "apple" and value > 5:**
```excel
=SUMIFS(C1:C10, B1:B10, "apple", C1:C10, ">5")
```

---

### 📉 AVERAGEIF()
Calculates the average of values in a range that meet a single condition.

**Syntax:**
```excel
=AVERAGEIF(range, criteria, [average_range])
```

**Example 1 — Average of values greater than 5:**
```excel
=AVERAGEIF(A1:A10, ">5")
```

**Example 2 — Average of values where text is "apple":**
```excel
=AVERAGEIF(B1:B10, "apple", C1:C10)
```

---

### 📉 AVERAGEIFS()
Calculates the average of values meeting multiple criteria.

**Syntax:**
```excel
=AVERAGEIFS(average_range, range1, criteria1, range2, criteria2, ...)
```

**Example 1 — Average of values between 5 and 10:**
```excel
=AVERAGEIFS(A1:A10, A1:A10, ">5", A1:A10, "<10")
```

**Example 2 — Average where text is "apple" and value > 5:**
```excel
=AVERAGEIFS(C1:C10, B1:B10, "apple", C1:C10, ">5")
```

---

### 🔡 LOWER()
Converts all letters in a text string to **lowercase**.

**Syntax:**
```excel
=LOWER(text)
```

**Example 1:**
```excel
=LOWER("EXCEL FORMULAS")
```
> Returns: `excel formulas`

**Example 2 — Convert a column of names to lowercase:**
```excel
=LOWER(A2)
```

---

### 🔠 UPPER()
Converts all letters in a text string to **UPPERCASE**.

**Syntax:**
```excel
=UPPER(text)
```

**Example 1:**
```excel
=UPPER("excel formulas")
```
> Returns: `EXCEL FORMULAS`

**Example 2:**
```excel
=UPPER(A2)
```

---

### 🅰️ PROPER()
Converts the **first letter of each word** to uppercase, all others to lowercase.

**Syntax:**
```excel
=PROPER(text)
```

**Example 1:**
```excel
=PROPER("eXCEL fORMULAS")
```
> Returns: `Excel Formulas`

**Example 2:**
```excel
=PROPER(A2)
```

---

### 🔻 MINIFS()
Returns the **minimum** value from a range that meets multiple criteria.

**Syntax:**
```excel
=MINIFS(min_range, range1, criteria1, range2, criteria2, ...)
```

**Example — Minimum order value for "West" region, "Product B", after Jan 2022:**
```excel
=MINIFS(D:D, A:A, "West", B:B, "Product B", C:C, ">"&DATE(2022,1,1))
```

---

### 🔺 MAXIFS()
Returns the **maximum** value from a range that meets multiple criteria.

**Syntax:**
```excel
=MAXIFS(max_range, range1, criteria1, range2, criteria2, ...)
```

**Example — Maximum sales for "East" region, "Product C", on or before Dec 31 2021:**
```excel
=MAXIFS(D:D, A:A, "East", B:B, "Product C", C:C, "<="&DATE(2021,12,31))
```

---

### 🔄 UNIQUE()
Returns a list of **unique values** from a range, removing duplicates.

**Syntax:**
```excel
=UNIQUE(array, [by_col], [exactly_once])
```

**Example 1 — Extract unique fruits from A2:A9:**
```excel
=UNIQUE(A2:A9)
```

**Example 2 — Extract unique values from two columns:**
```excel
=UNIQUE(A2:B9)
```

---

### ✂️ TRIM()
Removes all **leading and trailing spaces** from a text string.

**Syntax:**
```excel
=TRIM(text)
```

**Example 1 — Remove spaces from a string:**
```excel
=TRIM(A2)
```
> Input: `"  Hello World  "` → Output: `Hello World`

**Example 2 — Trim a range of names:**
```excel
=TRIM(A2)
```
*(drag down to apply to the full column)*

---

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

# DAY 4

## Chapter 7: Creating Charts and Graphics

### Steps to Create a Chart

1. **Select Data** — Organize with rows as categories and columns as values
2. **Insert Chart** — Click the **Insert** tab → choose chart type
3. **Customize Chart** — Add titles, labels, legends, adjust scales and colors

### Chart Types

| Chart Type | Best Used For | Example |
|------------|--------------|---------|
| **Column Chart** | Comparing values across categories | Comparing product sales |
| **Line Chart** | Showing trends over time | Stock prices over a year |
| **Pie Chart** | Showing percentage breakdown | Sales % per product |
| **Bar Chart** | Categories with long labels | Department performance |
| **Area Chart** | Trends over time with volume emphasis | Website traffic |
| **Scatter Chart** | Showing relationship between two variables | Ad spend vs sales |

---

## Chapter 8: Sparklines and Data Bars

### ⚡ Sparklines
Small charts embedded within a cell that visualize trends quickly.

**How to Insert:**
1. Select the cell where you want the sparkline
2. Go to **Insert** → **Sparklines** → **Line**
3. Choose the data range → Click **OK**

### 📊 Data Bars
Visual bars inside cells showing relative values.

**How to Apply:**
1. Select the range of cells
2. Go to **Home** → **Conditional Formatting** → **Data Bars**
3. Choose color/style → Click **OK**

> 💡 The longer the bar, the higher the value — great for quick visual comparison.

---

## Chapter 9: PivotTables and PivotCharts

### 📋 Creating a PivotTable

**Steps:**
1. **Select the data** (including headers)
2. Go to **Insert** → **PivotTable**
3. Choose the data source and where to place the PivotTable (**New Worksheet** or **Existing Worksheet**)
4. Click **OK** — a blank PivotTable is created
5. **Add fields** by dragging them in the PivotTable Fields pane:

| Area | Purpose | Example |
|------|---------|---------|
| **Rows** | Groups data by rows | Group by Product |
| **Columns** | Groups data by columns | Group by Month |
| **Values** | Summarizes data | Total Sales Amount |
| **Filters** | Filters the entire PivotTable | Filter by Region |

6. **Customize** — change layout, apply filters, change calculation types

---

### 📈 Creating a PivotChart

1. Create a PivotTable first (follow steps above)
2. Click any cell inside the PivotTable
3. Go to **Insert** → **PivotChart** → choose chart type
4. Customize using the **Chart Design** and **Format** tabs

**Example — Total Sales by Product Category and Region:**
1. Select dataset → **Insert** → **PivotTable** → **OK**
2. Drag **Product Category** to **Rows**, **Region** to **Columns**, **Sales Amount** to **Values**
3. Select any PivotTable cell → **Insert** → **PivotChart** → choose chart type

---

## Chapter 10: What-If Analysis

Explores different scenarios by changing input values to observe effects on outputs.

### Tools Available

| Tool | What It Does | Example Use Case |
|------|-------------|-----------------|
| **Goal Seek** | Finds the input needed to achieve a target output | What sales volume gives a 20% profit margin? |
| **Data Tables** | Creates a table of scenarios by varying 1 or 2 inputs | Revenue table for different price/quantity combos |
| **Scenario Manager** | Creates and compares multiple sets of input values | Comparing loan repayments at different interest rates |
| **Solver** | Finds optimal solution subject to constraints | Minimize cost while meeting demand constraints |

**How to Access:**
- **Goal Seek:** Data tab → What-If Analysis → Goal Seek
- **Data Tables:** Data tab → What-If Analysis → Data Table
- **Scenario Manager:** Data tab → What-If Analysis → Scenario Manager
- **Solver:** Data tab → Solver *(may need to enable in Add-Ins)*

---

## Chapter 11: Data Cleaning

The process of removing or correcting inaccurate, incomplete, or irrelevant data.

### Technique 1 — Removing Duplicate Data
**Data** tab → **Remove Duplicates** → select columns → **OK**

### Technique 2 — Removing Blank Rows
**Home** → **Find & Select** → **Go To Special** → **Blanks** → **OK** → Right-click → **Delete**

### Technique 3 — Correcting Spelling Errors
**Home** → **Find & Replace** → enter misspelled word in "Find" → correct spelling in "Replace" → **Replace All**

### Technique 4 — Converting Text to Numbers
Select range → Right-click → **Format Cells** → **Number** → **OK**

### Technique 5 — Removing Unwanted Characters
**Home** → **Find & Replace** → enter character in "Find what" → leave "Replace with" blank → **Replace All**

### Technique 6 — Handling Missing Data
```excel
=IF(ISBLANK(A2), "Default Value", A2)
```
> Checks if A2 is blank and replaces it with "Default Value" if true

### Technique 7 — Standardizing Text Case
```excel
=LOWER(A2)     -- all lowercase
=UPPER(A2)     -- ALL UPPERCASE
=PROPER(A2)    -- Title Case
=TRIM(A2)      -- Remove extra spaces
```

### Technique 8 — Correcting Spelling with Find & Replace
**Home** → **Find & Select** → **Replace** → Enter incorrect spelling → Enter correct spelling → **Replace All**

---

## 📌 Quick Reference Formula Sheet

| Function | Syntax | Purpose |
|----------|--------|---------|
| SUM | `=SUM(A1:A10)` | Total of a range |
| AVERAGE | `=AVERAGE(A1:A10)` | Mean of a range |
| COUNT | `=COUNT(A1:A10)` | Count of numbers |
| MAX | `=MAX(A1:A10)` | Highest value |
| MIN | `=MIN(A1:A10)` | Lowest value |
| CONCATENATE | `=CONCATENATE(A1," ",B1)` | Join text strings |
| LEFT | `=LEFT(A1, 5)` | First N characters |
| RIGHT | `=RIGHT(A1, 4)` | Last N characters |
| TRIM | `=TRIM(A1)` | Remove extra spaces |
| LOWER | `=LOWER(A1)` | lowercase text |
| UPPER | `=UPPER(A1)` | UPPERCASE text |
| PROPER | `=PROPER(A1)` | Title Case Text |
| IF | `=IF(A1>10,"Yes","No")` | Conditional logic |
| IFS | `=IFS(A1>90,"A",A1>80,"B")` | Multiple conditions |
| AND | `=AND(A1>5,B1<10)` | All conditions true? |
| OR | `=OR(A1=1,A1=2)` | Any condition true? |
| VLOOKUP | `=VLOOKUP("X",A1:C10,2,FALSE)` | Vertical lookup |
| XLOOKUP | `=XLOOKUP("X",A1:A10,B1:B10)` | Modern lookup |
| COUNTIF | `=COUNTIF(A1:A10,">5")` | Count with 1 criterion |
| COUNTIFS | `=COUNTIFS(A1:A10,">5",B1:B10,"Y")` | Count with multiple criteria |
| SUMIF | `=SUMIF(A1:A10,">5",B1:B10)` | Sum with 1 criterion |
| SUMIFS | `=SUMIFS(C1:C10,A1:A10,">5",B1:B10,"Y")` | Sum with multiple criteria |
| AVERAGEIF | `=AVERAGEIF(A1:A10,">5")` | Average with 1 criterion |
| AVERAGEIFS | `=AVERAGEIFS(C1:C10,A1:A10,">5")` | Average with multiple criteria |
| MINIFS | `=MINIFS(D:D,A:A,"West",B:B,"ProductB")` | Min with multiple criteria |
| MAXIFS | `=MAXIFS(D:D,A:A,"East",B:B,"ProductC")` | Max with multiple criteria |
| UNIQUE | `=UNIQUE(A2:A20)` | Remove duplicates |
| SWITCH | `=SWITCH(A1,"Yes",1,"No",0)` | Match cases |
| ISBLANK | `=ISBLANK(A1)` | Check if cell is empty |

---

*These notes were compiled for a 4-Day Excel Bootcamp. For more resources, visit [gatituportfolio.netlify.app/](https://gatituportfolio.netlify.app/)*
