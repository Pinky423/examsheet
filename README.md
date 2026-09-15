
📊 EXCEL SALES ANALYSIS & DASHBOARD 

📌 PROJECT OVERVIEW


This project is an Excel-based Sales Order Analysis and Interactive
Dashboard. It covers the complete workflow:

Data Import → Data Cleaning → Data Validation → Lookup Formulas →
Conditional Logic → Pivot Analysis → KPI Analysis → Charts → Interactive
Dashboard → Documentation

 
📂 WORKBOOK STRUCTURE

Sheet Purpose

Dashboard Interactive dashboard, KPIs, filters and charts
Data Cleaned master order dataset and helper/formula columns
Pivot Tables Category, region, payment and monthly summaries
Formula Checks Advanced formulas with examples and validation
Documentation Task-wise completion and workbook guide
Lists Lookup and Data Validation source lists


🧹 DATA CLEANING & PREPARATION


order_date standardized to YYYY-MM-DD

Markdown/separator row removed

Text values cleaned and trimmed

Numeric fields converted to proper numeric format

Dates converted to proper Excel dates

Consistent number formats applied

Master data organized as an Excel Table

Helper columns added for analysis

Main fields
order_id, order_date, customer_id, product_category, region,
quantity, unit_price, discount, payment_method, delivery_days,
customer_rating, revenue

🔎 DATA VALIDATION


Dashboard dropdowns are provided for: 1. Product Category 2.
Region 3. Payment Method

The source lists are maintained in the Lists sheet.

Purpose: controlled input, fewer spelling errors, consistent
selections and an interactive dashboard.

🔗 VLOOKUP
Purpose: retrieve Customer Code from Customer ID.

=VLOOKUP(C4,Lists!$A$4:$B$192,2,FALSE)
C4 = Customer ID

Lists!A:B = lookup table

2 = return column

FALSE = exact match

 🔗 INDEX + MATCH

 
Purpose: retrieve Category Code from Product Category.

=INDEX(Lists!$D$4:$D$7,MATCH(E4,Lists!$C$4:$C$7,0))
This provides a flexible lookup where lookup and return columns can be
independent.

📈 PIVOT ANALYSIS


Product Category
Orders

Quantity

Revenue

Average Rating

Region
Orders

Quantity

Revenue

Average Delivery Days

Payment Method
Orders

Revenue

Share of Orders

Monthly Trend
Month

Orders

Revenue

Quantity

🎯 KPI DASHBOARD


The Dashboard provides: - Total Revenue - Total Orders -
Quantity Sold - Average Customer Rating - Average Delivery
Days

The KPI calculations respond to the selected dashboard filters.

🎛️ INTERACTIVE FILTERS


Use the yellow dropdown cells for: - Product Category - Region - Payment
Method

Select All or a specific value. The KPI cards and supporting
analysis update through Excel formulas.

📊 DATA VISUALIZATION


Revenue by Category
Bar chart answering: Which category generates the most revenue?

Monthly Revenue Trend
Line chart answering: How does revenue change month by month?

Payment Method Mix
Doughnut chart answering: Which payment methods are most frequently
used?

📋 MASTER DATA TABLE


The Data sheet is the master source and includes helper columns: -
Month - Revenue Check - Delivery Status - Rating Band - Customer Code -
Category Code

The original business fields remain available for row-level analysis.

🎨 FORMATTING & HIGHLIGHTING


The workbook uses: - Clear section headers - Highlighted KPI cards -
Yellow input/dropdown cells - Conditional formatting for revenue -
Highlighting for delayed deliveries - Highlighting for excellent
ratings - Consistent date and number formats - Freeze panes for easier


💾 FINAL SUBMISSION


Recommended Excel filename:

YourName_ExcelDashboard.xlsx

Keep the Excel workbook and this README file together when submitting.

🏁 PROJECT SUMMARY


This project demonstrates a complete Excel data-analysis workflow using:

Excel Tables + Data Cleaning + Data Validation + Lookup Functions +
Conditional Logic + SUMIFS/COUNTIFS/AVERAGEIFS + Pivot Analysis + KPIs +
Charts + Interactive Dashboard

The Data sheet is the master source, Pivot Tables and Formula
Checks provide analytical support, and Dashboard presents the
final results in a user-friendly format.
