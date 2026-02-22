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
