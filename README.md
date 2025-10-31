##  About Me
**Power BI Vistualizations/Dashboard screenshots are at the end of the file**

Strategic Planning Analyst specializing in SIOP, forecasting, inventory optimization, and workflow automation. I eliminate manual business processes by automating workflows and designing end-to-end data reporting solutions and developing dashboards with Python, SQL, and Power BI, transforming large datasets into actionable insights. I thrive in dynamic environments and excel at building scalable solutions from the ground up.

# 📊 Inventory Analytics Portfolio (SIOP, Prioritization, Forecast Accuracy, Demand Forecasting)

This portfolio showcases six advanced reporting tools powered by Python, Pandas, Statistical models, and forecasting logic. All reports use dummy data and simulate real-world demand planning and inventory logic for supply chain operations. 
These reports are automated end-to-end using SQL Queries, SQL raw data tables from the MRP system, Python scripts with complex logic statements, and Power BI as the main workflow pipeline to extract, clean, consolidate, transform, and export data. They were designed to eliminate manual processes, drive cost savings of millions of dollars, and reduce human error, leveraging Python and SQL to enable large-scale data automation and support the company’s digital transformation strategy.

At my current company, I was spending hours every week manually consolidating data from multiple sources just to create a recurring report. It was extremely time-consuming and also increased the risk of errors that drained my energy before I really got to the analysis part. Colleagues in Supply Chain, Product Management, Sales and Operations were struggling with the same. Things I did to automate the data cleaning and consolidation process:
 
♦️ Identified repetitive Excel tasks (including VLOOKUPS, SUMIFS, TRIM, PivotTables, IFERROR, COUNTIF, MAX, MIN, among many others) and created a Python workflow using connections to raw SQL data tables and SQL Queries to consolidate data.  
♦️ Connected the data to a Power BI dashboard with a one-click refresh.  
♦️ Shared it across the team and scheduled the updates based on the report.  

This process for all reporting for SIOP resulted in saving:

✅ ~620+ hours a month of manual work for myself, the SIOP Manager, Product Managers, the Global Sales Operations Leader, and all the Supply Chain.

✅ Gave us more time to actually analyze the data and insights strategically rather than spending numerous hours putting files together to make a report.


The Python libraries used for these reports, to manipulate, transform, and analyze large datasets include:

- Pandas
- NumPy
- Matplotlib
- XlsxWriter (to format and export Excel report to Teams for Executives not familiar with Power BI features)
- Openpyxl
- For Demand Modeling: Statsmodels, Prophet, scikit-learn, scipy.stats
- Seaborn
- ThreadPoolExecutor (To speed up file reading process by reading multiple files simultaneously)

## 1. SIOP Time-Series & Capacity Automation (Power BI + Python/SQL): Single Source of Truth for Operations, Production and Planning

### Problem

Manual JDE/Excel/Salesforce workflows made it impossible to:

- Reconcile Demand/MRP/Work Orders against shift capacity by month and week.
- See Past Due load and material/line constraints at a glance.
- Run mass “what-if” analyses (thousands of SKUs) fast enough for real decisions.

### What I built

- A governed analytics engine + Power BI model that automated the manual Excel and JDE reporting workflow.

### Data/logic (Python + SQL)

- Ingests data from multiple JDE tables, plus Capacity & Material Supply from an Excel file. 
- Cleans, normalizes, and merges Q T → Type/Category, Planner → SIOP Valid, Alpha → Production Line.
- Adds Past Due classification (WO status ranges), SKU normalization, date hardening, and ASP overrides.
- Writes a Power BI–ready vertical file with strict column contracts and Excel formatting for audit.

### Key features (for stakeholders)

- Single source of truth for Operations/Planning/Production/Demand/Purchasing.
- Past Due surfaced explicitly; line/group filters for fast triage.
- Capacity realism: optional 2nd shift & Max overlays; work-day normalization.

### Results

- Became the company's standard dashboard for Operations.
- Over 60+ hours/week saved.
- Runs the entire pipeline and refresh Power BI in 3 minutes.

## 2. Inventory Prioritization Report Tool

Automates classification of backlog data into meaningful categories using rule-based logic consisting mainly of nested if statements. Needed to continue part two and appropriately categorize excess inventory by Buyer Name, Family, SKU, and others.

### Features
- Status Mapping, Percent Of Complete (POC) logic, Hold conditions
- Report Category with 8+ logic layers
- Formatted Excel output with visual styling

---

## 3. SIOP Inventory Burndown Tool (Anonymized Version)
Powered by [inventory_burndown_tool.py](./inventory_burndown_tool.py)

A time-phased inventory allocation report simulating excess-to-safety logic, POC allocations, and historical usage analysis. This tool automates inventory allocation, burn-down analysis, and reporting for supply chain and SIOP processes. It consolidates SQL data and multiple Excel reports data to generate actionable insights for planners, buyers, and executives. It's helped identify millions of dollars in Excess Buyer Inventory (purchases by Buyer Name) and Excess Finished Goods that were produced and don't have orders within a year timeframe. It helps Sales Executives sell excess inventory and helps Production Managers, Engineers, and others turn excess inventory into re-work without worrying about prioritizing inventory allocation. In addition to the final Power BI output, there is an individual output Excel file for those executives who are more familiar working in an Excel environment. All the calculations for the burndown are automatd by nested if logic statements in Python. There are no formulas in any cell for the calculations, reducing accidental click errors by 100%. 

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

## 4. Forecast Accuracy Tracker & Power BI Measures

Tracks MS Forecast vs Requested Sales and calculates accuracy KPIs.

### Features
- Auto-appends new forecast vs actuals data monthly
- DAX measures for WMAPE, WMAPE %, Forecast Accuracy %, Forecast Bias %
- Dynamic Date formatting and structure for Power BI use

### 📄 Outputs
- 📊 Power BI dashboards created for trend and KPI analysis
- [Forecast Accuracy KPI Dashboard Sample](Forecast_Accuracy_Report_PowerBI_Export.pdf)



--- Forecasting Models below use ML and Statistical Models that are used to supplement Forecasting Software capabilities as well as linking to Power BI for better data visualization. Data is already pre-processed by the use of other Python and SQL scripts and queries for big data manipulation. 

## 5. Forecast Weighted Adjustment Based on Historical Sales

Adjusts 2025 forecast based on 2024 historicals by Region and Alpha Name using dynamic month-level weights.

### Features
- Saves 20 hours weekly of manual Excel work
- Matches monthly trends using 2024 as baseline
- Preserves zero forecasts
- Automatically adjusts all 2025 columns in-place by Region > Family > Customer Name > SKU


## 6. Demand Plan Forecast with Best Statistical Model Fit & Override Scaling

Forecasts 2025 using Exponential Smoothing, Moving Average, ARIMA, and other statistical models. Applies override to meet global unit cap as agreed with Sales, and generates before/after visual trends.

### Features
- Fits statistical model per Region > Family > Customer Name > SKU
- Creates monthly forecast with override logic
- Outputs Excel workbook with two charts: unadjusted vs overridden
- Executes in under 2 minutes for thousands of SKUs

## 7. Buyouts Forecasting

Buy-out parts often exhibit sporadic, irregular demand, which makes them difficult to forecast using standard methods.
This script was developed to:

- Give extra weight to recent patterns while smoothing out historical noise
- Classify SKUs into business-relevant groups based on frequency, volatility, and recency
- Apply appropriate statistical models even for parts with minimal or intermittent demand

By improving buy-out forecasts, we can:

- **Reduce stock-outs** of long-lead or high-risk parts
- **Avoid excess inventory** and obsolescence exposure
- **Optimize working capital** across procurement and planning
- **Align forecasts** more tightly with SIOP processes

This Python script:

- Loads historical usage and lead time data from Excel
- Segments SKUs into 'High Run-Rate', 'True Buy-Out', 'Seasonal', 'Intermittent', etc.
- Applies ETS (for continuous demand) and TSB (for intermittent demand) models using `statsforecast`
- Calculates dynamic safety stock based on 12-month medians and lead time
- Caps forecast outliers to prevent inflated projections
- Outputs 'Individual', 'Overrides', and 'Adjusted' Excel sheets with dynamic formatting

Powered by [Buyouts_Forecast](./Buyouts_Business_Rules_Forecast.py) 

## 📸 Screenshots from Projects that required extensive backend and Power BI development, including the use of Vega-Lite to design custom visualizations

### Demand Plan Output Power BI Snapshots
![Demand Plan Value](<Demand_Plan_Value.png>)

### Demand Plan Output Power BI Snapshot by Region
![Demand_Plan_Units](<Demand_Plan_Units.png>)

### Time Series - Single source of truth for Demand/MRP/WO/Production Plan/Supply Plan/Capacity & Material Supply

### Monthly View Value
![Time_Series_Value](<Time_Series_Monthly_Value.png>)

### Monthly View Units
![Time_Series_Value](<Time_Series_Monthly_Units.png>)

### Weekly View Value (drill-down)
![Time_Series_Value](<Time_Series_Weekly_Value.png>)

### Weekly View Value
![Time_Series_Value](<Time_Series_Weekly_Value.png>)

### Weekly View Units by total working days 
![Time_Series_Value](<Time_Series_Workday_Units.png>)

### Forecast Accuracy KPI Snapshot
![Forecast Accuracy KPI snapshot](<Forecast_Accuracy_KPI.png>)

### Forecast Accuracy – Top Product Families
![Forecast Accuracy Top Product Families](<Forecast_Accuracy_Top_Product_Families.png>)

### Monthly Forecast Accuracy
![Montlhly Forecast Accuracy](<Forecast_Accuracy_KPI_by_Month.png>)

### Excess Saleable Finished Goods Inventory Power BI View
![Excess Saleable Inventory Power BI](<Excess_FG_Inventory_List.png>)

### Inventory Burn Down by Category
![Inventory Burn Down by Category](<Inventory_Burn_Down.png>)

### Buyer Inventory – Excess Purchases by Buyer Name
![Buyer Inventory](<Buyer_Excess_Inventory.png>)

