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
