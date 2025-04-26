## 👋 About Me

Bilingual Demand Planning Analyst with a strong background in SIOP forecasting, data automation, and advanced analytics. Skilled in Python, SQL, and Power BI, with a portfolio of tools that automate and optimize supply chain operations.

# 📊 Inventory Analytics Portfolio (SIOP, Prioritization, Forecast Accuracy, Adjustments)

This portfolio showcases six advanced Excel-based reporting tools powered by Python, Pandas, statistical models, and forecasting logic. All reports use dummy data and simulate real-world demand planning and inventory logic for supply chain operations.
These reports are automated end-to-end and were designed to eliminate manual processes, drive cost savings, and reduce human error, leveraging Python and SQL to enable large-scale data automation and support the company’s digital transformation strategy.

## 1. Inventory Prioritization Report Tool

Automates classification of backlog data into meaningful categories using rule-based logic consisting mainly of nested if statements. Needed to continue part two and appropriately categorize excess inventory by Buyer Name, Family, SKU, and others.

### Features
- Status Mapping, Percent Of Complete (POC) logic, Hold conditions
- Report Category with 8+ logic layers
- Formatted Excel output with visual styling

### Tools Used
Python (Pandas, OpenPyXL), Excel

---

## 2. SIOP Inventory Burndown Tool (Anonymized Version)
Powered by [inventory_burndown_tool.py](./inventory_burndown_tool.py)

A time-phased inventory allocation report simulating excess-to-safety logic, POC allocations, and historical usage analysis. This tool automates inventory allocation, burn-down analysis, and reporting for supply chain and SIOP processes. It consolidates multiple Excel reports to generate actionable insights for planners, buyers, and executives.


###  Features
- Saves 10 hours weekly of manual Excel work and publishes daily refreshable list to Power BI for Sales Team to reduce finished goods on hand inventory
- Monthly columns from PD through latest date
- Allocates supply across 9 SIOP categories
-  Consolidates inventory valuation, demand, and usage data
- Classifies inventory into burn-down categories (0-6 months, 1-2 years, etc.)
- Flags excess inventory based on average usage and safety stock
- Outputs Excel reports with clear formatting and prioritization
- Parallel file reading for speed and scalability
- Professional Excel formatting and multithreaded data loading

### Tools Used
Python (Pandas, OpenPyXL), Excel, Parallel processing with `ThreadPoolExecutor`

### 📄 Outputs
- [`📄 Download SIOP_Inventory_Report_Power BI (PDF)`](SIOP_Inventory_Report_PowerBI_Export.pdf)
- 📊 Multiple dashboards created in Power BI for internal use
- [SIOP Inventory Excel Output](SIOP_Inventory_Demo.xlsx)
  
## Use Case

Originally developed to support SIOP planning and reduce excess inventory across 5,000+ SKUs. This version uses dummy data to demonstrate how the tool enables:

- Inventory visibility by burn-down timeline
- Prioritization for rework and excess sell-through
- Alignment of supply and demand for executive decisions

## Disclaimer

This is a simulated version with anonymized paths and logic. No proprietary or company-sensitive information is included.

---
---

## 3. Forecast Accuracy Tracker & Power BI Measures

Tracks MS Forecast vs Requested Sales and calculates accuracy KPIs.

### Features
- Auto-appends new forecast vs actuals data monthly
- DAX measures for WMAPE, WMAPE %, Forecast Accuracy %, Forecast Bias %
- Dynamic Date formatting and structure for Power BI use

### Tools Used
Python (Pandas, OpenPyXL), Power BI

### 📄 Outputs
- 📊 Power BI dashboards created for trend and KPI analysis
- [Forecast Accuracy KPI Dashboard Sample](Forecast_Accuracy_Report_PowerBI_Export.pdf)



--- Forecasting Models below use ML and Statistical Models that are used to supplement Forecasting Software capabilities as well as linking to Power BI for better data visualization. Data is already pre-processed by the use of other Python and SQL scripts and queries for big data manipulation. 

## 4. Forecast Weighted Adjustment Based on Historical Sales

Adjusts 2025 forecast based on 2024 historicals by Region and Alpha Name using dynamic month-level weights.

### Features
- Saves 20 hours weekly of manual Excel work
- Matches monthly trends using 2024 as baseline
- Preserves zero forecasts
- Automatically adjusts all 2025 columns in-place by Region > Family > Customer Name > SKU

### Tools Used
Python (Pandas)


## 5. Demand Plan Forecast with Best Statistical Model Fit & Override Scaling

Forecasts 2025 using Exponential Smoothing, Moving Average, ARIMA, and other statistical models. Applies override to meet global unit cap as agreed with Sales, and generates before/after visual trends.

### Features
- Fits statistical model per Region > Family > Customer Name > SKU
- Creates monthly forecast with override logic
- Outputs Excel workbook with two charts: unadjusted vs overridden
- Executes in under 2 minutes for thousands of SKUs

### Tools Used
Python (Pandas, pmdarima, matplotlib, XlsxWriter)

## 6. Buyouts Forecast Adjustment Script

This Python script:
- Loads a wide-format history+forecast Excel file.
- Segments SKUs into continuous, intermittent, and buy-out.
- Runs ETS or TSB models and caps future values.
- Outputs `Overrides` and `Adjusted` sheets.

Powered by [Buyouts_Forecast_Adjuster](./Buyouts_Forecast_Adjuster.py)

## 📸 Screenshots

### Demand Plan Output Power BI Snapshot
![Demand Plan Output Power BI snapshot](<Demand_Plan.png>)

### Demand Plan Output Power BI Snapshot by Region
![Demand Plan Output Power BI snapshot by Region](<Demand_Plan_Region.png>)

### Forecast Accuracy KPI Snapshot
![Forecast Accuracy KPI snapshot](<Forecast_Accuracy_KPI.png>)

### Forecast Accuracy – Top Product Families
![Forecast Accuracy Top Product Families](<Forecast_Accuracy_Top_Product_Families.png>)

### Monthly Forecast Accuracy
![Montlhly Forecast Accuracy](<Forecast_Accuracy_KPI_by_Month.png>)

### Excess Saleable Inventory Power BI View
![Excess Saleable Inventory Power BI](<Excess_FG_Inventory_List.png>)

### Inventory Burn Down by Category
![Inventory Burn Down by Category](<Inventory_Burn_Down.png>)

### Buyer Inventory – Excess Purchases by Buyer
![Buyer Inventory](<Buyer_Excess_Inventory.png>)

