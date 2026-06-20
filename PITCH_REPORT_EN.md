# Sales Reporting System Pitch Report

## 1. Business Background

The company records daily customer orders on handwritten delivery notes. Staff then transfer each order into a separate Excel file, including products, quantities, and prices.

Current problems:

- Orders are stored separately, making daily totals difficult to see.
- Month-end consolidation requires significant manual work.
- Customer, product, and category performance is difficult to compare.
- Filename, date, and data-entry mistakes are easy to miss.
- Management lacks clear and timely operating data.

## 2. Proposed Solution

Install a Sales Reporting System on the company’s Windows 11 PC.

The system reads all daily order Excel files, validates their data, combines them, and automatically generates daily and monthly reports.

```text
Handwritten delivery note
→ Staff enters one order Excel file
→ Order file is saved in the correct date folder
→ System reads all order files
→ Daily and monthly reports are generated automatically
```

## 3. Standard Order File

Each Excel file represents one customer order.

Filename format:

```text
YYYYMMDD_CustomerName_Sequence.xlsx
```

Example:

```text
20260619_SeaviewRestaurant_001.xlsx
```

The order template contains:

- customer name;
- order date;
- delivery/order number;
- product category: beef, pork, fish, lamb, or sundries;
- product and meat cut;
- quantity sold;
- average purchase price;
- cheaper-cut or fine-cut selling price;
- total order quantity and revenue.

## 4. Nontechnical Staff Workflow

Desktop shortcut:

```text
Sales Report System
```

Daily process:

1. Click **Open Today’s Order Folder**.
2. Place today’s order Excel files into the folder.
3. Click **Generate Today’s Report**.
4. Click **Open Today’s Report**.

Monthly process:

1. Click **Generate This Month’s Report**.
2. Click **Open This Month’s Report**.

No PowerShell commands, programming, or complex Excel formulas are required.

## 5. Automation

- Automatically selects today’s order folder.
- Automatically creates folders for today and the next seven days.
- Automatically names daily reports: `DailyReport_YYYYMMDD.xlsx`.
- Automatically names monthly reports: `MonthlyReport_YYYYMM.xlsx`.
- Preserves previous reports instead of overwriting them on later dates.
- Remembers default order and report locations.
- Warns users when they select a folder for the wrong date.
- Warns users when an open Excel report prevents updating.

## 6. Data Validation

The system checks:

- whether filenames follow the required format;
- whether filename date matches the Excel date;
- whether Excel date matches the selected date folder;
- missing customer name, date, or order number;
- invalid or missing quantity;
- missing selling price;
- both selling-price columns being filled simultaneously;
- rows containing quantity or price but no product name;
- files containing no sold products.

Problems appear in the report’s **Issues** sheet for staff correction.

## 7. Daily Report

The daily report includes:

- total quantity sold;
- total revenue;
- customer totals;
- product totals;
- category totals;
- complete sales detail;
- data-entry issue list.

## 8. Monthly Report Analytics

The monthly report uses eight main sheets:

```text
Overview
Monthly Comparison
Customer Analysis
Product Analysis
Category Analysis
Daily Trend
Details
Issues
```

### Overview

- order count;
- customer count;
- total quantity sold;
- total revenue;
- average order value;
- average selling price;
- issue count.

### Customer Analysis

- revenue and purchased quantity by customer;
- customer share of revenue;
- top ten customers;
- current-month versus previous-month performance.

### Product Analysis

- quantity and revenue by product;
- average selling price;
- top ten products;
- product demand increases and decreases;
- current-month versus previous-month comparison.

### Category Analysis

- quantity and revenue for beef, pork, fish, lamb, and sundries;
- category share of revenue;
- current-month versus previous-month category comparison.

### Daily Trend

- daily order count;
- daily quantity sold;
- daily revenue.

## 9. Charts and Tables

The monthly report automatically creates five Excel charts:

- revenue share by category;
- daily revenue trend;
- top customer revenue;
- top product revenue;
- current-month versus previous-month comparison.

Main datasets are presented as sortable and filterable Excel tables.

## 10. Business Value

- Reduces daily and month-end consolidation work.
- Reduces date, filename, and data-entry errors.
- Gives management faster visibility into revenue and quantities.
- Identifies important customers and high-demand products.
- Monitors prices and revenue changes.
- Builds reliable historical data for future purchasing and inventory decisions.
- Preserves the company’s familiar Excel workflow, reducing training requirements.

## 11. Current Project Status

Completed:

- single-order Excel template;
- daily order folder structure;
- Excel reader and validation;
- daily report generation;
- monthly report generation;
- current-month versus previous-month comparisons;
- Excel tables and charts;
- nontechnical graphical interface;
- Windows desktop shortcut;
- remembered input and output locations;
- automatic date folders and report naming.

## 12. Recommended Next Phase

1. Pilot the system with real company orders.
2. Confirm all product names and categories.
3. Add duplicate-order detection and import history.
4. Add a central database.
5. Add user accounts and permissions.
6. Add automatic backup and recovery.
7. Package the system as a formal Windows installer.
8. Add demand forecasting and purchasing recommendations after several months of reliable data.

## 13. 60-Second Pitch

The company currently records orders on handwritten delivery notes and transfers each order into a separate Excel file. Because those files remain separate, management cannot quickly see daily revenue, total quantities, customer performance, or monthly trends.

The Sales Reporting System keeps the company’s familiar Excel workflow. Staff place each order file into the correct daily folder and click one button. The system automatically reads, validates, and combines every order into a daily report.

At month-end, it consolidates the entire month and presents customer, product, category, revenue, quantity, price, and previous-month comparisons. It also creates clear Excel charts automatically.

The interface is designed for nontechnical staff. It reduces manual work and errors while creating a reliable sales-data foundation for future forecasting, purchasing, and inventory management.
