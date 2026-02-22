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

## Day 4 Assignment — Charts, PivotTables & Data Cleaning

### Task 1: Creating Charts
Using your Sales Data:

1. Create a **Column Chart** showing Total Sales by Product
2. Create a **Pie Chart** showing the percentage of sales by Region
3. Create a **Line Chart** — add a Month column to your data and show sales trends over time
4. For each chart:
   - Add a chart title
   - Add axis labels
   - Change the chart colour scheme
   - Add data labels

### Task 2: Sparklines and Data Bars
1. Add a **Sparkline** (line type) to summarise each product's monthly sales trend in a single cell
2. Apply **Data Bars** to the Total Sales column to show relative performance visually

### Task 3: PivotTables and PivotCharts
Using your Sales Data:

1. Create a **PivotTable** that shows:
   - Total Sales by Product (rows) and Region (columns)
2. Add a **filter** to the PivotTable for Month
3. Change the value field to show **Average** instead of Sum
4. Create a **PivotChart** from your PivotTable (use a bar chart)
5. Add a slicer to make filtering interactive

### Task 4: What-If Analysis
1. Set up a simple profit model with: Revenue, Cost, and Profit (Revenue minus Cost)
2. Use **Goal Seek** to find what Revenue is needed to achieve a profit of $10,000
3. Create a **Data Table** showing how profit changes as revenue increases from $5,000 to $50,000 in steps of $5,000
4. Use **Scenario Manager** to create three scenarios: Best Case, Base Case, and Worst Case — each with different revenue and cost values

### Task 5: Data Cleaning
Download or create a messy dataset with the following issues intentionally built in:

1. Duplicate rows — use **Remove Duplicates** to clean them
2. Blank rows — use **Go To Special** to find and delete them
3. Inconsistent text case (e.g. "NAIROBI", "nairobi", "Nairobi") — use `=PROPER()` to standardise
4. Extra spaces in names — use `=TRIM()` to clean them
5. Numbers stored as text — convert them to actual numbers
6. Missing values — use `=IF(ISBLANK(A2), "Unknown", A2)` to fill them with a default value

---
