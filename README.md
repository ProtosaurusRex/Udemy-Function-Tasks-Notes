# Udemy-Function-Tasks-Notes
A collection of assignments I completed while taking a Python Udemy Course
Here's a clean, portfolio-ready `README.md` for your second project, this time with a stronger emphasis on the **backend skills** demonstrated—especially Excel automation, custom function creation, and programmatic data processing using Python. It's designed to showcase your ability to handle real-world tasks that mimic internal tooling or data infrastructure development.

---

# 💼 Backend Excel Automation – Maven Ski Shop Pricing & Tax Fixes

> **Project from:** [Python Foundations for Data Analysis – Udemy Course](https://www.udemy.com/course/python-foundations-for-data-analysis/)

## 📌 Project Overview

In this project, I built backend-style Python automations for a simulated retail company facing data consistency and international expansion challenges. The task involved using Python to:

1. **Fix missing sales tax entries** in an Excel order log.
2. **Create a dynamic currency converter** to support expansion into the UK and Japan.
3. **Write new calculated values directly to Excel**, modifying structured data at scale.

This project provided hands-on experience with **Python scripting**, **Excel data manipulation**, and **developing backend logic** to support frontend operations (like pricing pages or tax documentation).

---

## 🛠️ Technologies Used

| Tool / Library    | Purpose                                         |
| ----------------- | ----------------------------------------------- |
| `Python 3`        | Primary scripting language                      |
| `openpyxl`        | Programmatically read/write Excel files         |
| `Custom Modules`  | Used for tax calculations (`tax_calculator.py`) |
| `Function Design` | Created reusable logic for currency conversion  |
| `.xlsx` Format    | Manipulated structured retail datasets          |

---

## 📋 Business Task Summary

### 🔧 Task 1: Sales Tax Correction

> *"A customer reached out about a missing sales tax charge on their order. Please calculate the tax and update it in the Excel sheet (Row 10, Column D). Tax rate is 8%."*

✅ **Skills Applied:**

* Read specific cell data from Excel
* Applied domain logic using a custom `tax_calculator` module
* Wrote results back to Excel
* Ensured correct targeting of Excel rows/columns

📌 **Backend Code Highlights:**

```python
import openpyxl as xl
import tax_calculator as tc

wb = xl.load_workbook(filename='maven_ski_shop_data.xlsx')
order_info_sheet = wb['Orders_Info']

# Calculate sales tax on the value in cell D10
tc.tax_calculator(order_info_sheet['D10'].value, 0.08)
```

---

### 🌍 Task 2: International Currency Expansion

> *"We’re preparing for expansion into the EU, UK, and Japan. Please build a flexible currency converter and apply it to our item prices."*

✅ **Skills Applied:**

* Created a reusable **currency conversion function**
* Applied exchange rates using **loop-based data processing**
* Programmatically added **new columns** to Excel for GBP and JPY prices
* Preserved original data integrity while expanding structure

📌 **Backend Code Highlights:**

```python
def currency_converter(price, exchange_rate):
    return price * exchange_rate
```

```python
# GBP conversion
for i, cell in enumerate(items['G'], start=1):
    if i == 1:
        items[f'G{i}'] = 'GBP Price'
    else:
        items[f'G{i}'] = currency_converter(items[f'C{i}'].value, 0.76)

# JPY conversion
for i, cell in enumerate(items['H'], start=1):
    if i == 1:
        items[f'H{i}'] = 'JPY Price'
    else:
        items[f'H{i}'] = currency_converter(items[f'C{i}'].value, 123)

wb.save('maven_data_new_pricing.xlsx')
```

---

## 🧠 Backend Concepts Demonstrated

| Concept                               | Application                                                      |
| ------------------------------------- | ---------------------------------------------------------------- |
| **Data Layer Automation**             | Programmatic updates to Excel files (input/output)               |
| **Business Logic Implementation**     | Applied tax and currency logic in real-time                      |
| **Function Reusability**              | Modularized conversion logic for easy scaling                    |
| **Data Integrity & Targeted Updates** | Updated specific cells/columns without affecting other data      |
| **Internationalization Support**      | Introduced GBP/JPY pricing columns for UK/Japan expansion        |
| **Versioned Output Files**            | Saved transformed workbook as a new version to preserve original |

---

## 📁 File Structure

```
📁 maven_ski_shop_backend
├── maven_ski_shop_data.xlsx              # Original dataset
├── maven_data_new_pricing.xlsx           # Updated with tax + currency columns
├── tax_calculator.py                     # Custom module for tax logic
├── tax_and_currency_update.py            # Main script
└── README.md
```

---

## ✅ What I Learned

* How to **build backend processes** for spreadsheet-based systems
* Precise targeting of Excel cells and ranges with `openpyxl`
* Designing **general-purpose utility functions** (e.g., `currency_converter`)
* Structuring code for **real-world maintainability**
* Handling **multi-step data pipelines**: Read → Transform → Write

---

## 🔮 Possible Extensions

* Convert script to support **batch processing of customer tax issues**
* Integrate real-time currency API (e.g., **Fixer.io**) for dynamic conversion rates
* Add unit tests for core functions (`pytest`)
* Refactor for use with `pandas` for scalability and performance
* Deploy as a Flask API endpoint for internal tool usage

---

## 📬 Contact

Have feedback or want to collaborate on similar backend automation projects? Feel free to reach out!

