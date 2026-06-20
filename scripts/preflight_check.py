from __future__ import annotations

import tempfile
import time
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook
from pypdf import PdfReader

from read_daily_orders import (
    detect_duplicate_orders,
    generate_daily_report,
    generate_monthly_report,
    read_order_file,
)


def make_order(
    path: Path,
    order_date: str,
    customer: str,
    order_number: str,
    entries: list[tuple[int, str, str, Decimal, str, Decimal]],
) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "訂單輸入"
    sheet["B5"] = customer
    sheet["F5"] = datetime.strptime(order_date, "%Y%m%d")
    sheet["H5"] = order_number
    for row, product, category, quantity, price_column, price in entries:
        sheet.cell(row, 1, float(quantity))
        sheet.cell(row, 3, product)
        sheet.cell(row, 4, category)
        sheet.cell(row, 5, product)
        sheet[f"{price_column}{row}"] = float(price)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def find_summary_value(path: Path, label: str) -> Decimal:
    workbook = load_workbook(path, data_only=True)
    sheet = workbook["總覽"]
    for row in sheet.iter_rows():
        if row[0].value == label:
            return Decimal(str(row[1].value))
    raise AssertionError(f"Missing summary label: {label}")


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def run() -> None:
    started = time.perf_counter()
    with tempfile.TemporaryDirectory(prefix="sales-preflight-") as temp:
        root = Path(temp)
        orders = root / "Orders"
        current = orders / "20260618"
        previous = orders / "20260518"

        current_file = current / "20260618_測試客戶_001.xlsx"
        make_order(
            current_file,
            "20260618",
            "測試客戶",
            "001",
            [
                (9, "牛測試", "牛", Decimal("2"), "F", Decimal("10")),
                (53, "豬測試", "豬", Decimal("3"), "G", Decimal("20")),
                (84, "羊測試", "羊", Decimal("4"), "H", Decimal("30")),
                (89, "雜貨測試", "雜貨", Decimal("5"), "F", Decimal("4")),
            ],
        )
        make_order(
            previous / "20260518_測試客戶_001.xlsx",
            "20260518",
            "測試客戶",
            "001",
            [(9, "牛測試", "牛", Decimal("1"), "F", Decimal("10"))],
        )

        rows, issues = read_order_file(current_file, datetime(2026, 6, 18).date())
        require(not issues, f"Valid order produced issues: {issues}")
        require(sum((row.revenue for row in rows), Decimal("0")) == Decimal("220.00"), "F/G/H revenue calculation failed")
        require({row.price_type for row in rows} == {"平均售價", "平價切", "精修切"}, "Price-type detection failed")

        wrong_date_rows, wrong_date_issues = read_order_file(current_file, datetime(2026, 6, 19).date())
        require(wrong_date_rows == rows, "Date warning changed valid sales data")
        require(any("資料夾日期" in issue.message for issue in wrong_date_issues), "Wrong-folder-date warning missing")

        duplicate_file = current / "20260618_測試客戶_002.xlsx"
        make_order(
            duplicate_file,
            "20260618",
            "測試客戶",
            "001",
            [(9, "牛測試", "牛", Decimal("2"), "F", Decimal("10"))],
        )
        duplicate_rows, _ = read_order_file(duplicate_file)
        duplicate_issues = detect_duplicate_orders(rows + duplicate_rows)
        require(any("重複訂單編號" in issue.message for issue in duplicate_issues), "Duplicate-order warning missing")
        duplicate_file.unlink()

        daily_xlsx = root / "Daily Reports" / "每日報表_20260618.xlsx"
        daily_counts = generate_daily_report(current, daily_xlsx)
        require(daily_counts == (1, 4, 0), f"Unexpected daily counts: {daily_counts}")
        require(daily_xlsx.exists() and daily_xlsx.with_suffix(".pdf").exists(), "Daily Excel/PDF missing")
        require(find_summary_value(daily_xlsx, "總收入 HKD") == Decimal("220"), "Daily total revenue incorrect")

        monthly_xlsx = root / "Monthly Reports" / "月報表_202606.xlsx"
        monthly_counts = generate_monthly_report(orders, "202606", monthly_xlsx)
        require(monthly_counts == (1, 4, 0), f"Unexpected monthly counts: {monthly_counts}")
        require(monthly_xlsx.exists() and monthly_xlsx.with_suffix(".pdf").exists(), "Monthly Excel/PDF missing")
        require(find_summary_value(monthly_xlsx, "總收入 HKD") == Decimal("220"), "Monthly total revenue incorrect")
        comparison = load_workbook(monthly_xlsx, data_only=True)["月度比較"]
        revenue_comparison = next(
            row for row in comparison.iter_rows(min_row=2, values_only=True) if row[0] == "總收入 HKD"
        )
        require(Decimal(str(revenue_comparison[1])) == Decimal("220"), "Current-month comparison revenue incorrect")
        require(Decimal(str(revenue_comparison[2])) == Decimal("10"), "Previous-month comparison revenue incorrect")
        require(len(PdfReader(monthly_xlsx.with_suffix(".pdf")).pages) == 2, "Monthly PDF is not two pages")

    elapsed = time.perf_counter() - started
    print(f"PASS: calculations, warnings, duplicate detection, Excel/PDF reports ({elapsed:.2f}s)")


if __name__ == "__main__":
    run()
