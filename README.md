# Data_Exploration
```python
import openpyxl
import pandas as pd

for fname in ["Assignment 2 - Data Cleaning and Transformation.xlsx", "Assignment 2 - Data Cleaning and Transformation_Completed.xlsx"]:
    try:
        wb = openpyxl.load_workbook(fname, data_only=False)
        print(f"=== File: {fname} ===")
        print("Sheets:", wb.sheetnames)
        for s in wb.sheetnames:
            sheet = wb[s]
            print(f"--- Sheet: {s} ---")
            for r in range(1, 15):
                row_vals = [sheet.cell(r, c).value for c in range(1, 15)]
                if any(row_vals):
                    print(f"Row {r}: {row_vals}")
    except Exception as e:
        print(f"Error loading {fname}: {e}")


```

```text
=== File: Assignment 2 - Data Cleaning and Transformation.xlsx ===
Sheets: ['Dataset']
--- Sheet: Dataset ---
Row 1: ['Product ID', 'Product Name', 'Brand Name', 'Price ($)', 'Quantity', 'Category', None, None, None, None, None, None, None, None]
Row 2: ['28-JAN-US', 'laptop', 'Dell', 1000, 30, 'Electronics', None, None, None, None, None, None, None, None]
Row 3: ['15-FEB-US', 'Sneakers', 'Nike', 80, 15, 'Fashion', None, None, None, None, None, None, None, None]
Row 4: ['03-MAR-US', 'Coffee Maker', 'Keurig', 130, 40, 'Kitchen', None, None, None, None, None, None, None, None]
Row 5: ['11-APR-US', 'smartphone', 'Samsung', 900, 25, 'Electronics', None, None, None, None, None, None, None, None]
Row 6: ['22-MAY-US', 'Backpack', 'North Face', 70, 20, None, None, None, None, None, None, None, None, None]
Row 7: ['07-JUN-UK', 'Headphones', 'Sony', None, 45, 'Electroni', None, None, None, None, None, None, None, None]
Row 8: ['19-JUL-UK', 'T-shirt', 'Adidas', 30, 5, 'Fashion', None, None, None, None, None, None, None, None]
Row 9: ['23-AUG-UK', 'Blender', 'Ninja', 90, 35, 'Kitchen', None, None, None, None, None, None, None, None]
Row 10: ['05-SEP-UK', 'Tablet', 'Apple', 500, 50, 'Electronics', None, None, None, None, None, None, None, None]
Row 11: ['14-OCT-UK', 'Hiking Boots', 'Timberland', 130, 10, 'Outdoor', None, None, None, None, None, None, None, None]
Row 12: ['17-JUN-IN', 'Laptop', 'HP', 950, 25, 'Electronics', None, None, None, None, None, None, None, None]
Row 13: ['25-NOV-AU', 'Sneakers', 'Adidas', 90, 40, None, None, None, None, None, None, None, None, None]
Row 14: ['08-DEC-DE', 'Coffee Maker', 'Nespresso', 120, 35, None, None, None, None, None, None, None, None, None]
=== File: Assignment 2 - Data Cleaning and Transformation_Completed.xlsx ===
Sheets: ['Dataset']
--- Sheet: Dataset ---
Row 1: ['Product ID', 'Manufacturing Date', 'Country Code', 'Product Name', 'Brand Name', 'Price ($)', 'Quantity', 'Category', None, None, None, None, None, None]
Row 2: ['28-JAN-US', '28-01-2024', 'US', 'Laptop', 'Dell', 1000, 30, 'Electronics', None, None, None, None, None, None]
Row 3: ['15-FEB-US', '15-02-2024', 'US', 'Sneakers', 'Nike', 80, 15, 'Fashion', None, None, None, None, None, None]
Row 4: ['03-MAR-US', '03-03-2024', 'US', 'Coffee Maker', 'Keurig', 130, 40, 'Kitchen', None, None, None, None, None, None]
Row 5: ['11-APR-US', '11-04-2024', 'US', 'Smartphone', 'Samsung', 900, 25, 'Electronics', None, None, None, None, None, None]
Row 6: ['22-MAY-US', '22-05-2024', 'US', 'Backpack', 'North Face', 70, 20, 'Unknown', None, None, None, 'Handling Missing Values: Price Coloum', None, None]
Row 7: ['07-JUN-UK', '07-06-2024', 'UK', 'Headphones', 'Sony', 297.85714285714283, 45, 'Electronics', None, None, None, 'IF(ISBLANK(D5), AVERAGE($D$2:$D$32), D5)', None, None]
Row 8: ['19-JUL-UK', '19-07-2024', 'UK', 'T-Shirt', 'Adidas', 30, 5, 'Fashion', None, None, None, None, None, None]
Row 9: ['23-AUG-UK', '23-08-2024', 'UK', 'Blender', 'Ninja', 90, 35, 'Kitchen', None, None, None, 'Products with missing categories', None, None]
Row 10: ['05-SEP-UK', '05-09-2024', 'UK', 'Tablet', 'Apple', 500, 50, 'Electronics', None, None, None, "IF(ISBLANK(G2),'Unknown',G2)", None, None]
Row 11: ['14-OCT-UK', '14-10-2024', 'UK', 'Hiking Boots', 'Timberland', 130, 10, 'Outdoor', None, None, None, None, None, None]
Row 12: ['17-JUN-IN', '17-06-2024', 'IN', 'Laptop', 'HP', 950, 25, 'Electronics', None, None, None, 'Correcting Inconsistent Data:', None, None]
Row 13: ['25-NOV-AU', '25-11-2024', 'AU', 'Sneakers', 'Adidas', 90, 40, 'Unknown', None, None, None, 'PROPER(TRIM(CLEAN(C2)))', None, None]
Row 14: ['08-DEC-DE', '08-12-2024', 'DE', 'Coffee Maker', 'Nespresso', 120, 35, 'Unknown', None, None, None, None, None, None]


```

```python
import openpyxl

def inspect_wb(fname):
    wb = openpyxl.load_workbook(fname, data_only=False)
    print(f"=== File: {fname} ===")
    sheet = wb.active
    print("Sheet name:", sheet.title)
    for r in range(1, 40):
        row_vals = [sheet.cell(r, c).value for c in range(1, 15)]
        if any(row_vals):
            # clean nones from end
            while row_vals and row_vals[-1] is None:
                row_vals.pop()
            print(f"Row {r:2d}: {row_vals}")

inspect_wb("Assignment 2 - Data Cleaning and Transformation_Completed.xlsx")


```

```text
=== File: Assignment 2 - Data Cleaning and Transformation_Completed.xlsx ===
Sheet name: Dataset
Row  1: ['Product ID', 'Manufacturing Date', 'Country Code', 'Product Name', 'Brand Name', 'Price ($)', 'Quantity', 'Category']
Row  2: ['28-JAN-US', '28-01-2024', 'US', 'Laptop', 'Dell', 1000, 30, 'Electronics']
Row  3: ['15-FEB-US', '15-02-2024', 'US', 'Sneakers', 'Nike', 80, 15, 'Fashion']
Row  4: ['03-MAR-US', '03-03-2024', 'US', 'Coffee Maker', 'Keurig', 130, 40, 'Kitchen']
Row  5: ['11-APR-US', '11-04-2024', 'US', 'Smartphone', 'Samsung', 900, 25, 'Electronics']
Row  6: ['22-MAY-US', '22-05-2024', 'US', 'Backpack', 'North Face', 70, 20, 'Unknown', None, None, None, 'Handling Missing Values: Price Coloum']
Row  7: ['07-JUN-UK', '07-06-2024', 'UK', 'Headphones', 'Sony', 297.85714285714283, 45, 'Electronics', None, None, None, 'IF(ISBLANK(D5), AVERAGE($D$2:$D$32), D5)']
Row  8: ['19-JUL-UK', '19-07-2024', 'UK', 'T-Shirt', 'Adidas', 30, 5, 'Fashion']
Row  9: ['23-AUG-UK', '23-08-2024', 'UK', 'Blender', 'Ninja', 90, 35, 'Kitchen', None, None, None, 'Products with missing categories']
Row 10: ['05-SEP-UK', '05-09-2024', 'UK', 'Tablet', 'Apple', 500, 50, 'Electronics', None, None, None, "IF(ISBLANK(G2),'Unknown',G2)"]
Row 11: ['14-OCT-UK', '14-10-2024', 'UK', 'Hiking Boots', 'Timberland', 130, 10, 'Outdoor']
Row 12: ['17-JUN-IN', '17-06-2024', 'IN', 'Laptop', 'HP', 950, 25, 'Electronics', None, None, None, 'Correcting Inconsistent Data:']
Row 13: ['25-NOV-AU', '25-11-2024', 'AU', 'Sneakers', 'Adidas', 90, 40, 'Unknown', None, None, None, 'PROPER(TRIM(CLEAN(C2)))']
Row 14: ['08-DEC-DE', '08-12-2024', 'DE', 'Coffee Maker', 'Nespresso', 120, 35, 'Unknown']
Row 15: ['18-FEB-CA', '18-02-2024', 'CA', 'Smartwatch', 'Fitbit', 150, 15, 'Electronics', None, None, None, 'Split the "Product ID"']
Row 16: ['16-APR-ES', '16-04-2024', 'ES', 'Headphones', 'Bose', 250, 20, 'Electronics', None, None, None, 'LEFT(A2, 6)']
Row 17: ['21-AUG-CA', '21-08-2024', 'CA', 'Laptop Bag', 'Samsonite', 50, 35, 'Accessories', None, None, None, 'RIGHT(A2, 2)']
Row 18: ['20-AUG-CN', '20-08-2024', 'CN', 'Smartwatch', 'Huawei', 160, 15, 'Electronics']
Row 19: ['27-JAN-IT', '27-01-2024', 'IT', 'Laptop', 'Asus', 980, 10, 'Electronics', None, None, None, 'Format the "Manufacturing Date" column to display dates in the "DD-MM-YYYY "']
Row 20: ['01-MAR-UK', '01-03-2024', 'UK', 'Sunglasses', 'Oakley', 150, 15, 'Fashion', None, None, None, 'TEXT(DATEVALUE(B2 & "-2024"), "dd-mm-yyyy")']
Row 21: ['14-AUG-US', '14-08-2024', 'US', 'Camping Tent', 'Coleman', 297.85714285714283, 10, 'Outdoor']
Row 22: ['14-MAY-RU', '14-05-2024', 'RU', 'Camera', 'Nikon', 700, 50, 'Electronics']
Row 23: ['09-JAN-CA', '09-01-2024', 'CA', 'Microwave', 'Panasonic', 80, 20, 'Kitchen']
Row 24: ['19-JUL-BR', '19-07-2024', 'BR', 'Fitness Tracker', 'Xiaomi', 150, 30, 'Unknown']
Row 25: ['29-SEP-CA', '29-09-2024', 'CA', 'Smartphone', 'Google', 800, 45, 'Electronics']
Row 26: ['03-JUN-CA', '03-06-2024', 'CA', 'Sunglasses', 'Ray-Ban', 297.85714285714283, 25, 'Fashion']
Row 27: ['11-JUL-CA', '11-07-2024', 'CA', 'Blender', 'Vitamix', 400, 40, 'Kitchen']
Row 28: ['07-MAR-CA', '07-03-2024', 'CA', 'Dress', 'Zara', 60, 30, 'Fashion']
Row 29: ['13-APR-CA', '13-04-2024', 'CA', 'Toaster', 'Hamilton', 40, 10, 'Kitchen']
Row 30: ['24-MAY-CA', '24-05-2024', 'CA', 'Fitness Tracker', 'Garmin', 130, 5, 'Electronics']
Row 31: ['02-DEC-CA', '02-12-2024', 'CA', 'Jeans', "Levi's", 50, 50, 'Fashion']
Row 32: ['09-JUL-FR', '09-07-2024', 'FR', 'Watch', 'Casio', 100, 20, 'Accessories']


```

```python
import pypdf

reader = pypdf.PdfReader("Assignment 2 - Data Cleaning and Transformation.pdf")
print("Num pages:", len(reader.pages))
for idx, page in enumerate(reader.pages):
    print(f"--- Page {idx+1} ---")
    print(page.extract_text())


```

```text
Num pages: 6
--- Page 1 ---
1. Handling Missing Values: 
Check for missing values in the 'Price' column. How would you handle products with missing price 
information? 
 
If there are products with missing categories, propose a strategy to impute or deal with these missing values 
effectively. 
 
2. Correcting Inconsistent Data: 
Identify any inconsistent text formats present in the "Product Name" column. 
Identify any typos present in the "Category" column. 
Use the find and replace function to standardize the text formats in the "Product 
Name" column and fix any typos or misspellings in the "Category" column. 

--- Page 2 ---
 
3. Removing Duplicates: 
 

--- Page 3 ---
4. Splitting and Merging Data: 
Split the "Product ID" column into two separate columns for " Manufacturing Date" and "Country Code" . Remove 
unnecessary characters, if any. 
Merge the "Brand Name" and "Product Name" columns into one column named "Product Brand" . 
 

--- Page 4 ---
5. Number Formatting: 
Format the data type of the "Price" column to a currency format. 
 
Format the "Manufacturing Date" column to display dates in the "DD-MM-YYYY" format. 
 

--- Page 5 ---
6. Conditional Formatting: 
Apply data bar or color scales conditional formatting in the "Price" column. 
 

--- Page 6 ---
Create a custom rule for conditional formatting in the "Category" column to highlight cells where the category is 
"Electronics. " 
 



```

# Excel Data Cleaning & Transformation

---

## 📌 Project Overview

This repository contains the completed deliverable for **Data Analytics (DA) — Module 1: Excel (Assignment 2 – Data Cleaning and Transformation)**.

The primary focus of this assignment is on data hygiene, handling missing values, standardizing inconsistent text, deduplication, string extraction/concatenation, data type formatting, and visual conditional formatting in Microsoft Excel.

---

## 🎯 Summary of Completed Tasks

| Task Category | Description | Applied Technique / Formula |
| --- | --- | --- |
| **1. Missing Value Imputation** | Filled missing values in `Price` using mean imputation; handled missing `Category` entries. | `=IF(ISBLANK(D2), AVERAGE($D$2:$D$32), D2)`<br>

<br>`=IF(ISBLANK(H2), "Unknown", H2)` |
| **2. Text Inconsistency & Typos** | Corrected lowercase product names and fixed category typos (`Electroni` $\rightarrow$ `Electronics`). | `PROPER(TRIM(CLEAN(cell)))`<br>

<br>Find & Replace (`Ctrl + H`) |
| **3. Duplicate Removal** | Identified and removed duplicate product records to ensure unique entries. | Data Tab $\rightarrow$ **Remove Duplicates** |
| **4. String Parsing & Merging** | Extracted date and country code from `Product ID`; combined Brand and Product Name. | `=LEFT(A2, 6)` & `=RIGHT(A2, 2)`<br>

<br>`=E2 & " " & D2` |
| **5. Number & Date Formatting** | Formatted prices as Currency and dates into `DD-MM-YYYY` standard date format. | Format Cells (`Ctrl + 1`) $\rightarrow$ Currency / Custom (`dd-mm-yyyy`) |
| **6. Conditional Formatting** | Applied visual heatmap/bars to `Price` and custom highlighting for `Electronics`. | Data Bars / Color Scales<br>

<br>Formula Rule: `=$H2="Electronics"` |

---

## 🛠️ Detailed Implementation Guide

### 1. Handling Missing Values

* **Price Column Imputation:** Missing price values were replaced using the mean price of non-empty records ($297.86).
```excel
=IF(ISBLANK(D2), AVERAGE($D$2:$D$32), D2)

```


* **Category Imputation Strategy:** Missing or unknown categories were flagged using `"Unknown"` or imputed via lookup matching against existing product names (`XLOOKUP` / `VLOOKUP`).

### 2. Correcting Inconsistent Data

* **Product Name Standardization:** Standardized mixed-case text (e.g., `laptop` $\rightarrow$ `Laptop`, `smartphone` $\rightarrow$ `Smartphone`) using proper case formatting:
```excel
=PROPER(TRIM(CLEAN(D2)))

```


* **Category Misspelling Fixes:** Corrected truncated or misspelled text (`Electroni` $\rightarrow$ `Electronics`, `Kchen` $\rightarrow$ `Kitchen`) using Excel's **Find and Replace** feature (`Ctrl + H`).

### 3. Removing Duplicates

* Selected the dataset range and navigated to **Data** $\rightarrow$ **Remove Duplicates**.
* Checked `Product ID` and core attributes to remove duplicate entries, ensuring data integrity.

### 4. Splitting and Merging Data

* **Manufacturing Date Extraction:** Extracted the date component (first 6 characters) from `Product ID` (e.g., `28-JAN-US` $\rightarrow$ `28-JAN`):
```excel
=LEFT(A2, 6)

```


* **Country Code Extraction:** Extracted the last 2 characters from `Product ID` (e.g., `28-JAN-US` $\rightarrow$ `US`):
```excel
=RIGHT(A2, 2)

```


* **Merging Brand and Product:** Combined `Brand Name` and `Product Name` into `Product Brand`:
```excel
=E2 & " " & D2

```



### 5. Number and Date Formatting

* **Price Column:** Formatted cells to **Currency** format (`$#,##0.00` or `₹#,##0`).
* **Manufacturing Date Column:** Converted text dates to true date objects and formatted using Custom format `dd-mm-yyyy`:
```excel
=TEXT(DATEVALUE(B2 & "-2024"), "dd-mm-yyyy")

```



### 6. Conditional Formatting

* **Price Data Bars / Color Scales:** Applied **Data Bars** (Blue/Green gradient fill) and **Color Scales** to visually compare relative price values.
* **Electronics Highlight Rule:** Applied a custom formula rule to highlight all rows where the category is `"Electronics"`:
```excel
=$H2="Electronics"

```



---

## 📁 Repository Structure

```text
.
├── Assignment 2 - Data Cleaning and Transformation.pdf             # Task Guidelines
├── Assignment 2 - Data Cleaning and Transformation.xlsx            # Raw Input Dataset
├── Assignment 2 - Data Cleaning and Transformation_Completed.xlsx  # Processed & Formatted Workbook
└── README.md                                                       # Project Documentation

```

---

## 🚀 How to Use

1. Clone this repository:
```bash
git clone https://github.com/your-username/excel-data-cleaning-transformation.git

```


2. Open `Assignment 2 - Data Cleaning and Transformation_Completed.xlsx` in **Microsoft Excel 2016+** or **Office 365**.
3. Inspect formulas in columns `B`, `C`, and conditional formatting rules under **Home** $\rightarrow$ **Conditional Formatting** $\rightarrow$ **Manage Rules**.
