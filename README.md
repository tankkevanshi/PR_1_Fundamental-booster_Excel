# 📈 Fundamental Booster: End-to-End Excel Data Modeling & Analytics

> **An interactive spreadsheet portfolio demonstrating dynamic lookups, conditional aggregations, automated grading logic, and string/date transformations across multi-domain datasets.**

---

## 📌 Project Summary

This project serves as a practical benchmark for essential-to-intermediate Microsoft Excel capabilities (`PR_1_Fundamental_Booster_Excel.xlsx`). Designed around three real-world business and academic scenarios—**Student Performance**, **Regional Sales**, and **Workforce Analytics**—the workbook eliminates manual data processing through scalable, error-handled formulas.

---

## 🧭 Workbook Architecture

| Tab | Domain | Records | Core Objective |
| :--- | :--- | :---: | :--- |
| **`Project Instructions`** | Documentation | — | Outlines project scope, formatting standards, and formula syntax references. |
| **`Students Grade`** | Education | 20 | Automates score aggregation, letter-grade classification, and high-scorer filtering. |
| **`Sales Data`** | Retail / B2B | 20 | Calculates tiered volume discounts, net revenue, and dynamic regional KPIs. |
| **`Employee Data`** | Human Resources | 20 | Standardizes text records, tracks employee tenure, and executes multi-criteria lookups. |

---

## ⚙️ Formula & Logic Breakdown

### 🎓 1. `Students Grade`
* **Score & Grade Automation:** Computes total scores (`SUM`) and rounded averages (`ROUND`, `AVERAGE`), feeding into a nested `IF` statement to assign letter grades (`A` ≥ 90, `B` ≥ 80, `C` ≥ 65, `D` < 65).
* **Dual-Subject Excellence Flag:** Uses `IF(AND(...))` to flag students scoring above 80 in both Math and Science.
* **Dynamic Querying:** Integrates `COUNTIFS` and `AVERAGEIF` for subject pass-rate metrics, `VLOOKUP` for ID-based student retrieval, and `FILTER` to dynamically isolate top performers (>80%).

### 💰 2. `Sales Data`
* **Tiered Discount Logic:** Evaluates transaction size via nested `IF` formulas to apply `15%` (≥ $30,000), `10%` (≥ $15,000), or `5%` discounts, followed by net revenue and `ROUND` calculations.
* **Targeted Aggregations:** Uses `SUMIFS` and `AVERAGEIFS` to slice revenue by product category (e.g., *Laptop*, *Mouse*) and territory (*West*, *South*, *East*).
* **Flexible Lookups & Dynamic Ranges:** Combines `VLOOKUP`, `INDEX/MATCH` wrapped in `IFERROR`, positional `MATCH`, `INDIRECT` for string-referenced range totals, and `OFFSET` to sum the top 5 transactions dynamically.

### 👥 3. `Employee Data`
* **Data Cleaning & Parsing:** Extracts first names from full strings using `LEFT` and `FIND`, while enforcing case consistency across names and departments via `UPPER` and `LOWER`.
* **Live Tenure Tracking:** Calculates exact years and total days of service against `TODAY()` using `DATEDIF`, safeguarded by `IFERROR`.
* **Workforce Analytics:** Executes reverse and exact lookups (`INDEX/MATCH`) by Employee ID and Salary, alongside `COUNTIFS` and `AVERAGEIFS` for departmental headcount and payroll benchmarking.

---

## 🖥️ Quick Start Guide

1. **Open the File:** Launch `PR_1_Fundamental_Booster_Excel.xlsx` in **Microsoft Excel (Microsoft 365 / Excel 2021+)** for full dynamic array support (`FILTER`).
2. **Test Interactive Cells:** Update the highlighted lookup inputs—such as **Search Student ID** (`I25`), **Enter Sales ID** (`F30`), **Salesperson** (`F33`), or **Employee ID** (`K27`)—to see real-time recalculations.
3. **Inspect Formulas:** Toggle formula view (`Ctrl + ~` on Windows / `Cmd + ~` on Mac) to audit cell referencing and nested syntax.

---

## 💡 Core Competencies Showcased

* **Referencing & Auditing:** Relative vs. absolute cell locking and `IFERROR` fault tolerance.
* **Conditional Logic:** Nested `IF`, `AND`, `COUNTIFS`, `SUMIFS`, and `AVERAGEIFS`.
* **Lookup & Reference Functions:** `VLOOKUP`, `INDEX/MATCH`, `MATCH`, `INDIRECT`, `OFFSET`, and `FILTER`.
* **Text, Math & Date Functions:** `LEFT`, `FIND`, `UPPER`, `LOWER`, `ROUND`, `DATEDIF`, and `TODAY`.
