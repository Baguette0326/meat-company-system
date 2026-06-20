# Excel Monthly Sales Reporting System

## 1. Current Scope

The company process is:

```text
Multiple pink 發貨單 each day
        ↓
Staff input each order into the system or Excel import format
        ↓
Software saves all order rows into a central database
        ↓
Software generates daily reports and monthly analysis
```

This project will **not** read the pink handwritten sheets for now. It assumes
staff manually input the pink-sheet details correctly.

## 2. Confirmed Company Decisions

| Topic | Confirmed decision |
|---|---|
| Source file | One Excel workbook per order |
| Date basis | Order date |
| Customer | One customer per workbook |
| `數目` column | Ignore |
| Sold quantity field | New column named `售出数量` |
| `HKD/KG` meaning | Correct label is `HKD/KG` |
| Column F | Average buy-in price |
| Column G | Cheaper meat-cut selling price |
| Column H | Finer meat-cut selling price |
| Multiple price columns per row | Not allowed |
| Blank product rows | Ignore |
| Decimals | Allowed for quantities and prices |
| Returns/cancellations/free items | Not supported for now |
| Main report metrics | Both quantity and revenue |
| Product categories | Beef, pig, chicken, fish, other |
| Comparisons | Current month vs previous month |
| Trends | Customer trends and product demand changes |
| Report language | Traditional Chinese |
| Exports | Excel and PDF |
| Company PC | Windows 11 |
| Users | Several people |
| Backup | Internal PC storage |

## 3. Important Open Issue

Resolved: the sold quantity will be entered in a new column named `售出数量`.
The existing `數目` column is ignored.

The spelling should be standardized everywhere as `售出数量`. If the workbook
uses Traditional Chinese-only labels later, use `售出數量` consistently instead.
The importer can support both labels, but the template should prefer one.

## 4. Excel Import Assumptions

The importer will read each completed workbook as one order.

| Workbook area | Meaning |
|---|---|
| Customer field near row 3 | Customer name |
| Date field near row 4 | Order date |
| Column C | Product name |
| Column D | Existing product/category text |
| Column E | Cut/part |
| Column F | Average buy-in price |
| Column G | Cheaper cut selling price |
| Column H | Finer cut selling price |
| New column `售出数量` | Sold quantity |

A product row is included when `售出数量` contains a positive number.

Current price rule:

- exactly one of columns G or H should contain the selling price for a sold row;
- column F is stored for margin/reference analysis, not treated as sales revenue;
- line revenue is calculated as `售出数量 × selling price`;
- if both G and H are filled, the row is rejected for review;
- if neither G nor H is filled for a sold item, the row is rejected for review.

## 5. Product Categorization

Each product must map to one reporting category:

- Beef
- Pig
- Chicken
- Fish
- Other

The system should keep a product master list:

| Field | Purpose |
|---|---|
| Product name | Name from the workbook |
| Category | Beef/pig/chicken/fish/other |
| Default cut type | Optional |
| Active | Whether this product is still used |

When the importer sees a new product name, it should ask staff to assign a
category once. Future imports reuse that category.

## 6. Validation Rules

Reject or flag an imported workbook when:

- customer is missing;
- order date is missing or invalid;
- `售出数量` is missing, zero, or negative for a sold row;
- a sold row has no product name;
- a sold row has both G and H filled;
- a sold row has neither G nor H filled;
- selling price is zero or negative;
- decimal values are not valid numbers;
- product category is unknown;
- the same file was already imported;
- a likely duplicate order is detected.

The system should ignore:

- blank rows;
- repeated header rows;
- product rows with no sales entry;
- `數目`.

## 7. Duplicate Protection

The exact invoice/order-number system will be decided later. Until then, the
system should detect duplicates using:

```text
File fingerprint + customer + order date + sold product rows + selling prices
```

Each import records:

- original filename;
- file fingerprint;
- import date and time;
- user who imported it;
- accepted row count;
- rejected row count.

## 8. Monthly Report

The monthly report will be in Traditional Chinese and export to both Excel and
PDF.

### Headline Summary

- monthly revenue;
- number of orders;
- number of active customers;
- average order value;
- previous-month revenue comparison;
- previous-month order-count comparison.

### Product Summary

Grouped by product and category:

- total revenue;
- sold quantity, if a quantity field is confirmed;
- average selling price;
- cheaper-cut vs finer-cut sales split;
- previous-month change;
- demand increase/decrease flag.

### Customer Summary

Grouped by customer:

- monthly revenue;
- order count;
- average order value;
- top products bought;
- previous-month change.

### Category Summary

Grouped by:

- beef;
- pig;
- chicken;
- fish;
- other.

Metrics:

- revenue;
- quantity, if available;
- percentage of total sales;
- previous-month change.

### Price And Margin View

Because column F is the average buy-in price, the system can estimate margin:

```text
Estimated margin per kg = selling price - average buy-in price
Estimated margin % = estimated margin / selling price
```

This should be labelled as an estimate unless the company confirms column F is
the exact cost for that order.

## 9. Company PC Deployment

The software should run on the company Windows 11 PC, not from this development
folder.

Recommended deployment:

- desktop shortcut;
- local web interface in the browser;
- local database on the company PC;
- support for several staff users;
- automatic backup to internal PC storage;
- administrator-only restore function;
- Excel/PDF report export folder;
- no programming tools required for staff.

Because several people will use the system, the first version should run as a
small local server on the company PC. Staff on the same network can connect to
it through a browser if the company permits internal-network access.

## 10. Daily Report

The daily report is for operational visibility after staff finish inputting the
day's 發貨單.

### Daily Summary

- total daily revenue;
- total number of orders;
- number of customers served;
- total sold quantity by product;
- total revenue by product;
- total revenue by category;
- cheaper-cut vs finer-cut revenue;
- estimated gross margin using column F as average buy-in price.

### Daily Product Table

| Product | Category | 售出数量 | Revenue | Avg Selling Price | Cut Type |
|---|---|---:|---:|---:|---|

### Daily Customer Table

| Customer | Orders | Revenue | Top Products |
|---|---:|---:|---|

### Daily Exceptions

- missing customer;
- missing order date;
- product without category;
- row with both G and H filled;
- row with no selling price;
- duplicate-looking order.

## 11. Monthly Report

The monthly report is for database usage and deeper analysis:

- current month vs previous month;
- customer trends;
- product demand changes;
- category performance;
- estimated margin trends;
- top increasing and decreasing products;
- customer concentration;
- exportable Excel and PDF reports.

## 12. Implementation Phases

### Phase 1: Finalize Excel Contract

- add the `售出数量` column to the Excel template;
- confirm exact customer-name cell;
- confirm exact date cell;
- confirm whether G and H are the only selling-price columns;
- create generated sample workbooks for testing.

### Phase 2: Importer

- read Excel workbooks;
- extract customer, order date, products, prices, and `售出数量`;
- classify products into beef/pig/chicken/fish/other;
- validate rows;
- detect duplicates;
- save accepted rows to the database.

### Phase 3: Daily Reports

- Traditional Chinese daily dashboard;
- daily product quantity and revenue report;
- daily customer report;
- daily category report;
- daily exceptions report.

### Phase 4: Monthly Reports

- Traditional Chinese dashboard;
- monthly revenue summary;
- product/category/customer summaries;
- current vs previous month comparison;
- customer trend report;
- product demand change report;
- Excel export;
- PDF export.

### Phase 5: Company PC Package

- package for Windows 11;
- create desktop shortcut;
- configure local database and backups;
- test with several user accounts;
- produce a short Traditional Chinese user guide.
