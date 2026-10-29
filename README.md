# Final-Project-Excel

# 📊 Customer Transactions Analytics Dashboard

## 📑 Table of Contents

1. [Overview](#overview)
2. [Dataset](#dataset)
3. [Workbook Structure](#workbook-structure)
4. [Key Features](#key-features)
5. [Dashboard KPIs](#dashboard-kpis)
6. [Sample Formulas](#sample-formulas)
7. [Insights & Storytelling](#insights--storytelling)
8. [How to Use](#how-to-use)
9. [Author](#author)

---

## 🎯 Overview

This project delivers an end-to-end analytics solution in **Microsoft Excel**. It analyses customer transactions, uncovers business insights, and presents them through visualizations and an interactive dashboard.

**Goals**

- Clean and structure customer transaction data
- Apply date, lookup, dynamic-array and statistical functions
- Run What-If analysis and linear regression
- Build an interactive KPI dashboard
- Turn numbers into a clear business story

## 🗂️ Dataset

Source sheet: **`Final Project Dataset`** (200+ customer transaction records).

**📈 Live snapshot** (auto-generated from the dataset):

| Metric | Value |
|---|---|
| Records | 210 |
| Total Revenue | 227,203.41 |
| Date Range | 11 Apr 2024 to 19 Jun 2024 |
| Unique Customers | 48 |
| Most Sold Product | Bookshelf |
| Top Region (Revenue) | Central |
| Top Customer | Kevin Scott |

**Schema**

| Column | Description | Type |
|---|---|---|
| `Transaction_ID` | Unique ID of each transaction | Text |
| `Date` | Transaction date | Date |
| `Customer_ID` | Unique customer identifier | Text |
| `Customer_Name` | Full name of customer | Text |
| `Product_ID` | Unique product identifier | Text |
| `Product_Name` | Name of product purchased | Text |
| `Category` | Product category (Electronics, Appliances, Furniture...) | Text |
| `Quantity` | Units purchased | Number |
| `Unit_Price` | Price per unit | Currency |
| `Payment_Method` | Cash / Credit Card / Debit Card / PayPal | Text |
| `Region` | Sales region (North, South, East, West, Central) | Text |
| `Customer_Segment` | Basic / Standard / Premium | Text |

## 🧱 Workbook Structure

| Sheet | Purpose |
|---|---|
| **Raw Data** | Original 200+ customer transaction records (untouched) |
| **Analysis** | Formulas: date functions, FILTER, GroupBy, lookups, TEXT abbreviations |
| **What-If** | Scenario Manager, Goal Seek and Regression output |
| **Visualizations** | Pivot tables and pivot charts (Bar, Line, Pie) |
| **Dashboard** | KPI cards, slicers, timelines and icon-set indicators |
| **Report** | Final insights and storytelling summary |

## ✨ Key Features

| Topic | Tools / Functions | Used For |
|---|---|---|
| Date & Time Functions | `TODAY`, `NOW`, `DATEDIF`, `EOMONTH` | Customer tenure, transaction age, month-end reporting |
| Dynamic Arrays | `FILTER` | Multi-value returns, e.g. all orders for a region or customer |
| Conditional Formatting | Icon sets and arrows | Visual KPI indicators and trend arrows |
| Timestamp | `NOW` + iterative calculation | Auto-stamping when data is entered or changed |
| What-If Analysis | Scenario Manager, Goal Seek | Best/worst-case revenue scenarios and target solving |
| Linear Regression | Data Analysis ToolPak | Quantity / price vs revenue relationship |
| GroupBy Concepts | `SUMIFS`, `UNIQUE`, `SORTBY` | Identify high-value customers |
| Dashboard | Pivot Tables, Pivot Charts, Slicers, Timelines | Interactive, filterable dashboard |
| Frequency Analysis | `COUNTIF`, `INDEX/MATCH`, `MODE` | Most frequently purchased product / repeat names |
| List Comparison | `XLOOKUP`/`MATCH`, `IF` | Extract matching names into a new column |
| Abbreviations | `TEXT`, `LEFT`, `UPPER` | Create short codes from names and categories |

## 📌 Dashboard KPIs

`Total Revenue` `Top Customers` `Most Sold Products` `Regional Sales`


The dashboard combines **Pivot Tables**, **Pivot Charts** (Bar, Line, Pie), **Slicers**, **Timelines** and **conditional-formatting KPI indicators**.

## 🧮 Sample Formulas

```excel
Revenue per row            =H2*I2
Total revenue              =SUM(M:M)
Months since purchase      =DATEDIF(B2,TODAY(),"m")
Month-end date             =EOMONTH(B2,0)
Orders for a region        =FILTER(A:L,K:K="North")
Customer total spend       =SUMIFS(M:M,C:C,C2)
Most sold product          =INDEX(F:F,MODE(MATCH(F2:F250,F2:F250,0)))
Month label                =TEXT(B2,"mmm-yy")
Timestamp                  =IF(A2<>"",NOW(),"")
```

> Adjust column letters to match your workbook layout.

## 📖 Insights & Storytelling

Summarise findings in the **Report** sheet:

1. **What happened?** Revenue trend over time
2. **Who drives it?** High-value customers and segments
3. **What sells?** Top products and categories
4. **Where?** Regional performance
5. **What next?** Recommendations from What-If and regression results


---
## 🎬 Project Demonstration
  <a href="https://drive.google.com/file/d/1mres0xBKH-RSmjPlqbQXHzA8gdPhoIt7/view?usp=sharing">
    <img src="https://img.shields.io/badge/🎬%20Project%20Video-success?style=for-the-badge">
  </a>
</p>

---
