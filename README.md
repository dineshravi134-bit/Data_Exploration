# Excel Data Cleaning & Transformation

A beginner-friendly data cleaning project completed for **Module 1 - Assignment 2 (Data Analytics)**. 

This project demonstrates how to clean, transform, and format raw product data in Microsoft Excel using formulas and built-in features.

---

## 🛠️ Tasks & Formulas Summary

| # | Task | Solution / Formula Used |
|---|---|---|
| **1** | **Fill Missing Prices** | `=IF(ISBLANK(D2), AVERAGE($D$2:$D$32), D2)` |
| **2** | **Fix Missing Category** | Assigned `"Unknown"` or mapped using `XLOOKUP` |
| **3** | **Fix Text / Typos** | Used `PROPER()` & **Find & Replace** (`Ctrl + H`) |
| **4** | **Remove Duplicates** | **Data** tab $\rightarrow$ **Remove Duplicates** |
| **5** | **Extract Date** | `=LEFT(A2, 6)` |
| **6** | **Extract Country Code** | `=RIGHT(A2, 2)` |
| **7** | **Merge Brand & Product** | `=E2 & " " & D2` |
| **8** | **Format Date & Currency** | Custom Format `dd-mm-yyyy` & Currency (`$`) |
| **9** | **Highlight "Electronics"** | Conditional Formatting Formula: `=$H2="Electronics"` |

---
