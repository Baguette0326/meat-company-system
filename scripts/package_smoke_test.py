from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook


def main() -> int:
    root = Path.cwd()
    scripts = root / "scripts"
    sys.path.insert(0, str(scripts))

    from master_data import load_master_products  # noqa: PLC0415
    from read_daily_orders import generate_daily_report  # noqa: PLC0415

    products = load_master_products(root / "Master Data" / "master-data.xlsx")
    if not products:
        raise SystemExit("No master products found")
    product = next(iter(products.values()))

    orders = root / "Orders" / "20260713"
    orders.mkdir(parents=True, exist_ok=True)
    order_path = orders / "20260713_測試客戶_001.xlsx"

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "訂單輸入"
    sheet["B5"] = "測試客戶"
    sheet["C5"] = "正常"
    sheet["F5"] = datetime(2026, 7, 13)
    sheet["H5"] = "001"
    sheet["A9"] = 2
    sheet["C9"] = product.name
    sheet["D9"] = product.category
    sheet["E9"] = product.part
    sheet["F9"] = 10
    workbook.save(order_path)

    output = root / "Daily Reports" / "每日報表_20260713.xlsx"
    counts = generate_daily_report(orders, output)
    if counts[0] != 1 or counts[1] != 1:
        raise SystemExit(f"Unexpected counts: {counts}")
    if not output.exists() or not output.with_suffix(".pdf").exists():
        raise SystemExit("Output report missing")

    print(f"PASS: packaged daily report smoke test {counts}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
