# Daily Order Reader

## Purpose

Reads a folder of order Excel files and creates one daily report.

Input files should use:

```text
YYYYMMDD_客戶名_序號.xlsx
```

Example:

```text
20260618_海景酒家_001.xlsx
20260618_大發餐廳_001.xlsx
20260618_大發餐廳_002.xlsx
```

## Run

```powershell
python '.\scripts\read_daily_orders.py' `
  '.\Orders\20260618' `
  --output '.\Daily Reports\每日報表_20260618.xlsx'
```

## One-Click App

Double-click:

```text
run_daily_report_app.bat
```

Then:

1. Click `選擇資料夾`.
2. Pick folder containing one day's order files.
3. Confirm output path.
4. Click `產生 / 更新報表`.
5. Click `開啟報表`.

For monthly report:

```text
Click 產生本月月報
```

Output:

```text
Monthly Reports\月報表_YYYYMM.xlsx
```

If report is already open in Excel, close it first before updating.

Each report action creates Excel and PDF output. Daily PDF uses one page; monthly PDF uses two pages.

```text
Daily Reports\每日報表_YYYYMMDD.xlsx
Daily Reports\每日報表_YYYYMMDD.pdf

Monthly Reports\月報表_YYYYMM.xlsx
Monthly Reports\月報表_YYYYMM.pdf
```

Use app buttons `打開今日 PDF` and `打開本月 PDF` to view them.

The app remembers:

- order base folder;
- last export folder.

On startup, the app auto-selects today's folder:

```text
Orders\YYYYMMDD
```

It also creates today plus the next 7 day folders automatically:

```text
Orders\20260618
Orders\20260619
Orders\20260620
...
```

Report filename is automatic:

```text
每日報表_YYYYMMDD.xlsx
```

Example:

```text
每日報表_20260618.xlsx
```

The date comes from the selected folder name first, then from order filenames.

The reader warns if an order file date or Excel date does not match the selected folder date.

Duplicate detection warns when:

- two files use the same date, customer, and delivery/order number;
- two files have the same customer, date, products, quantities, and prices.

Monthly report reads all day folders under `Orders` matching that month.

## Output Tabs

- `Summary`
- `By Customer`
- `By Product`
- `By Category`
- `Sold Rows`
- `Issues`

## Validation

The reader flags:

- filename not matching `YYYYMMDD_客戶名_序號.xlsx`;
- filename date/customer mismatch against the workbook;
- missing customer;
- missing or invalid date;
- missing order number;
- quantity without selling price;
- both selling price columns filled;
- selling price less than or equal to zero;
- entry row with quantity/price but no product name.

## Current Template Contract

The reader expects one sheet named `訂單輸入` or `每日銷售表`.

Important cells:

| Cell | Meaning |
|---|---|
| `B5` | Customer name |
| `F5` | Order date |
| `H5` | Delivery/order number |

Product rows use:

| Column | Meaning |
|---|---|
| `A` | 售出数量 |
| `C` | 品名 |
| `D` | 分類 |
| `E` | 部位 |
| `F` | Average buy-in price |
| `G` | Cheap-cut selling price |
| `H` | Fine-cut selling price |
