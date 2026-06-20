from __future__ import annotations

import argparse
import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


PRICE_TYPES = {
    "average": "平均售價",
    "cheap": "平價切",
    "fine": "精修切",
}

FILENAME_RE = re.compile(r"^(?P<date>\d{8})_(?P<customer>.+)_(?P<seq>\d{3})\.xlsx$", re.IGNORECASE)


@dataclass(frozen=True)
class OrderFileMeta:
    path: Path
    filename_date: date | None
    filename_customer: str | None
    sequence: str | None


@dataclass(frozen=True)
class SaleRow:
    source_file: str
    row_number: int
    order_date: date | None
    customer: str
    order_number: str
    category: str
    product: str
    part: str
    quantity: Decimal
    unit_price: Decimal
    price_type: str
    buy_in_price: Decimal | None
    revenue: Decimal
    estimated_margin: Decimal | None


@dataclass(frozen=True)
class ImportIssue:
    source_file: str
    row_number: int | None
    severity: str
    message: str


def to_decimal(value) -> Decimal | None:
    if value is None or value == "":
        return None
    try:
        return Decimal(str(value).strip())
    except (InvalidOperation, AttributeError):
        return None


def money(value: Decimal | None) -> Decimal | None:
    if value is None:
        return None
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def parse_filename(path: Path) -> OrderFileMeta:
    match = FILENAME_RE.match(path.name)
    if not match:
        return OrderFileMeta(path, None, None, None)

    raw_date = match.group("date")
    try:
        parsed_date = datetime.strptime(raw_date, "%Y%m%d").date()
    except ValueError:
        parsed_date = None

    return OrderFileMeta(
        path=path,
        filename_date=parsed_date,
        filename_customer=match.group("customer"),
        sequence=match.group("seq"),
    )


def normalize_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value is None or value == "":
        return None
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def cell_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def get_sheet(workbook):
    if "訂單輸入" in workbook.sheetnames:
        return workbook["訂單輸入"]
    if "每日銷售表" in workbook.sheetnames:
        return workbook["每日銷售表"]
    return workbook[workbook.sheetnames[0]]


def read_header_fields(sheet):
    customer = cell_text(sheet["B5"].value)
    order_date = normalize_date(sheet["F5"].value)
    order_number = cell_text(sheet["H5"].value)

    # Some users may enter the values beside the labels on row 5, but keep a
    # fallback search in case the layout is slightly shifted.
    if not customer or not order_date or not order_number:
        for row in sheet.iter_rows(min_row=1, max_row=8, min_col=1, max_col=8):
            values = [cell_text(cell.value) for cell in row]
            for index, value in enumerate(values):
                next_value = values[index + 1] if index + 1 < len(values) else ""
                if not customer and value.startswith("客戶名稱"):
                    customer = next_value
                if not order_date and value.startswith("日期"):
                    order_date = normalize_date(next_value)
                if not order_number and value.startswith("發貨單編號"):
                    order_number = next_value

    return customer, order_date, order_number


def is_section_or_header(row_values: list[str]) -> bool:
    first = row_values[0]
    product = row_values[2]
    if first in {"牛", "豬", "鱼", "魚", "羊", "雜貨", "杂货"}:
        return True
    if first == "售出数量" or product == "品名":
        return True
    if first in {"總結", "訂單總結", "每日總結"}:
        return True
    return False


def iter_workbooks(folder: Path) -> Iterable[Path]:
    for path in sorted(folder.glob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        yield path


def folder_date(folder: Path) -> date | None:
    if not re.fullmatch(r"\d{8}", folder.name):
        return None
    try:
        return datetime.strptime(folder.name, "%Y%m%d").date()
    except ValueError:
        return None


def read_order_file(path: Path, expected_folder_date: date | None = None) -> tuple[list[SaleRow], list[ImportIssue]]:
    issues: list[ImportIssue] = []
    rows: list[SaleRow] = []
    meta = parse_filename(path)

    if meta.filename_date is None or meta.filename_customer is None:
        issues.append(
            ImportIssue(
                path.name,
                None,
                "warning",
                "檔名應使用 YYYYMMDD_客戶名_序號.xlsx",
            )
        )

    try:
        workbook = load_workbook(path, data_only=False)
    except Exception as exc:  # noqa: BLE001 - report workbook-level import failure.
        return [], [ImportIssue(path.name, None, "錯誤", f"無法開啟 Excel 檔案: {exc}")]

    sheet = get_sheet(workbook)
    customer, order_date, order_number = read_header_fields(sheet)

    if not customer:
        issues.append(ImportIssue(path.name, None, "錯誤", "B5 缺少客戶名稱"))
    if not order_date:
        issues.append(ImportIssue(path.name, None, "錯誤", "F5 缺少日期或日期格式無效"))
    if not order_number:
        order_number = meta.sequence or ""
        issues.append(ImportIssue(path.name, None, "警告", "H5 缺少發貨單編號"))

    if meta.filename_date and order_date and meta.filename_date != order_date:
        issues.append(
            ImportIssue(
                path.name,
                None,
                "警告",
                f"檔名日期 {meta.filename_date.isoformat()} 與 Excel 日期 {order_date.isoformat()} 不一致",
            )
        )
    if expected_folder_date and meta.filename_date and meta.filename_date != expected_folder_date:
        issues.append(
            ImportIssue(
                path.name,
                None,
                "警告",
                f"檔名日期 {meta.filename_date.isoformat()} 與資料夾日期 {expected_folder_date.isoformat()} 不一致",
            )
        )
    if expected_folder_date and order_date and order_date != expected_folder_date:
        issues.append(
            ImportIssue(
                path.name,
                None,
                "警告",
                f"Excel 日期 {order_date.isoformat()} 與資料夾日期 {expected_folder_date.isoformat()} 不一致",
            )
        )
    if meta.filename_customer and customer and meta.filename_customer != customer:
        issues.append(
            ImportIssue(
                path.name,
                None,
                "警告",
                f"檔名客戶「{meta.filename_customer}」與 Excel 客戶「{customer}」不一致",
            )
        )

    # Product entry rows start after the top order header. Rows above 7 contain
    # customer/date/order fields, which must not be interpreted as item prices.
    for row_number in range(7, sheet.max_row + 1):
        values = [sheet.cell(row_number, col).value for col in range(1, 9)]
        text_values = [cell_text(value) for value in values]
        if not any(text_values):
            continue
        if is_section_or_header(text_values):
            continue

        quantity = to_decimal(values[0])
        product = text_values[2]
        category = text_values[3]
        part = text_values[4]
        # F 欄是所有分類適用的平均售價。
        average_price = to_decimal(values[5])
        buy_in_price = None
        cheap_price = to_decimal(values[6])
        fine_price = to_decimal(values[7])
        has_any_entry = any(value is not None for value in [quantity, values[5], cheap_price, fine_price])

        if not product and has_any_entry:
            issues.append(ImportIssue(path.name, row_number, "錯誤", "此行有数量/價錢，但沒有貨品名稱"))
            continue
        if not product:
            continue
        if quantity is None and not average_price and not cheap_price and not fine_price:
            continue

        if quantity is None:
            issues.append(ImportIssue(path.name, row_number, "錯誤", f"{product}: 缺少售出数量"))
            continue
        if quantity <= 0:
            issues.append(ImportIssue(path.name, row_number, "錯誤", f"{product}: 售出数量必須大於 0"))
            continue
        entered_prices = [price for price in (average_price, cheap_price, fine_price) if price is not None]
        if len(entered_prices) > 1:
            issues.append(ImportIssue(path.name, row_number, "錯誤", f"{product}: 只可填寫一種售價"))
            continue
        if not entered_prices:
            issues.append(ImportIssue(path.name, row_number, "錯誤", f"{product}: 缺少售價"))
            continue

        if average_price is not None:
            unit_price = average_price
            price_type = PRICE_TYPES["average"]
        elif cheap_price is not None:
            unit_price = cheap_price
            price_type = PRICE_TYPES["cheap"]
        else:
            unit_price = fine_price
            price_type = PRICE_TYPES["fine"]
        if unit_price is None or unit_price <= 0:
            issues.append(ImportIssue(path.name, row_number, "錯誤", f"{product}: 售價必須大於 0"))
            continue

        revenue = money(quantity * unit_price)
        estimated_margin = money(quantity * (unit_price - buy_in_price)) if buy_in_price is not None else None

        rows.append(
            SaleRow(
                source_file=path.name,
                row_number=row_number,
                order_date=order_date,
                customer=customer,
                order_number=order_number,
                category=category,
                product=product,
                part=part,
                quantity=quantity,
                unit_price=unit_price,
                price_type=price_type,
                buy_in_price=buy_in_price,
                revenue=revenue or Decimal("0"),
                estimated_margin=estimated_margin,
            )
        )

    if not rows:
        issues.append(ImportIssue(path.name, None, "警告", "沒有找到已售出貨品"))

    return rows, issues


def detect_duplicate_orders(sale_rows: list[SaleRow]) -> list[ImportIssue]:
    issues: list[ImportIssue] = []
    rows_by_file: dict[str, list[SaleRow]] = defaultdict(list)
    for row in sale_rows:
        rows_by_file[row.source_file].append(row)

    seen_identity: dict[tuple, str] = {}
    seen_content: dict[tuple, str] = {}

    for source_file, rows in sorted(rows_by_file.items()):
        if not rows:
            continue
        first = rows[0]
        identity = (first.order_date, first.customer, first.order_number)
        if first.order_number:
            previous_file = seen_identity.get(identity)
            if previous_file:
                issues.append(
                    ImportIssue(
                        source_file,
                        None,
                        "警告",
                        f"重複訂單編號：與 {previous_file} 使用相同日期、客戶及發貨單編號",
                    )
                )
            else:
                seen_identity[identity] = source_file

        content_lines = tuple(
            sorted(
                (
                    row.category,
                    row.product,
                    str(row.quantity),
                    str(row.unit_price),
                    row.price_type,
                )
                for row in rows
            )
        )
        content_signature = (first.order_date, first.customer, content_lines)
        previous_content_file = seen_content.get(content_signature)
        if previous_content_file and previous_content_file != seen_identity.get(identity):
            issues.append(
                ImportIssue(
                    source_file,
                    None,
                    "警告",
                    f"疑似重複訂單內容：與 {previous_content_file} 的客戶、日期、貨品、数量及售價相同",
                )
            )
        elif not previous_content_file:
            seen_content[content_signature] = source_file

    return issues


def safe_table_name(title: str) -> str:
    # Excel table names are strict; use ASCII regardless of sheet title.
    suffix = abs(hash(title)) % 1_000_000
    return f"Table_{suffix}"


def add_sheet(workbook: Workbook, title: str, headers: list[str], data: list[list], make_table: bool = True):
    sheet = workbook.create_sheet(title)
    sheet.append(headers)
    for row in data:
        sheet.append(row)
    style_header(sheet)
    if make_table and headers:
        end_row = max(2, len(data) + 1)
        end_col = len(headers)
        table = Table(displayName=safe_table_name(title), ref=f"A1:{get_column_letter(end_col)}{end_row}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        sheet.add_table(table)
    autosize(sheet)
    return sheet


def style_header(sheet):
    fill = PatternFill("solid", fgColor="115E59")
    font = Font(bold=True, color="FFFFFF")
    for cell in sheet[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.row_dimensions[1].height = 34
    sheet.freeze_panes = "A2"


def autosize(sheet):
    for column in sheet.columns:
        width = 10
        letter = get_column_letter(column[0].column)
        for cell in column:
            value = "" if cell.value is None else str(cell.value)
            display_width = sum(2 if unicodedata.east_asian_width(char) in {"W", "F", "A"} else 1 for char in value)
            width = max(width, min(display_width + 3, 48))
        sheet.column_dimensions[letter].width = width


def append_sheet_section(target, source, section_title: str) -> None:
    start_row = target.max_row + 3
    target.cell(start_row, 1, section_title)
    target.cell(start_row, 1).font = Font(bold=True, size=14, color="115E59")
    for source_row in source.iter_rows(values_only=True):
        target.append(list(source_row))
    header_row = start_row + 1
    fill = PatternFill("solid", fgColor="115E59")
    font = Font(bold=True, color="FFFFFF")
    for cell in target[header_row]:
        if cell.value is not None:
            cell.fill = fill
            cell.font = font
    autosize(target)


def consolidate_monthly_sheets(workbook: Workbook) -> None:
    customer_sheet = workbook["客戶月報"]
    product_sheet = workbook["貨品月報"]
    category_sheet = workbook["分類月報"]
    issue_sheet = workbook["問題"]

    append_sheet_section(customer_sheet, workbook["客戶比較"], "客戶本月與上月比較")
    append_sheet_section(product_sheet, workbook["貨品比較"], "貨品本月與上月比較")
    append_sheet_section(category_sheet, workbook["分類比較"], "分類本月與上月比較")
    append_sheet_section(issue_sheet, workbook["問題摘要"], "問題摘要")

    customer_sheet.title = "客戶分析"
    product_sheet.title = "貨品分析"
    category_sheet.title = "分類分析"

    for sheet_name in ["客戶比較", "貨品比較", "分類比較", "十大客戶", "十大貨品", "問題摘要"]:
        workbook[sheet_name].sheet_state = "hidden"


def add_monthly_charts(
    workbook: Workbook,
    summary_sheet,
    category_sheet,
    trend_sheet,
    top_customer_sheet,
    top_product_sheet,
    comparison_sheet=None,
):
    dashboard = summary_sheet

    if category_sheet.max_row >= 2:
        pie = PieChart()
        labels = Reference(category_sheet, min_col=1, min_row=2, max_row=category_sheet.max_row)
        data = Reference(category_sheet, min_col=4, min_row=1, max_row=category_sheet.max_row)
        pie.add_data(data, titles_from_data=True)
        pie.set_categories(labels)
        pie.title = "分類收入佔比"
        pie.height = 8
        pie.width = 10
        dashboard.add_chart(pie, "D2")

    if trend_sheet.max_row >= 2:
        line = LineChart()
        data = Reference(trend_sheet, min_col=5, min_row=1, max_row=trend_sheet.max_row)
        labels = Reference(trend_sheet, min_col=1, min_row=2, max_row=trend_sheet.max_row)
        line.add_data(data, titles_from_data=True)
        line.set_categories(labels)
        line.title = "每日收入走勢"
        line.y_axis.title = "收入 HKD"
        line.x_axis.title = "日期"
        line.height = 8
        line.width = 16
        dashboard.add_chart(line, "N2")

    if top_customer_sheet.max_row >= 2:
        bar = BarChart()
        data = Reference(top_customer_sheet, min_col=2, min_row=1, max_row=top_customer_sheet.max_row)
        labels = Reference(top_customer_sheet, min_col=1, min_row=2, max_row=top_customer_sheet.max_row)
        bar.add_data(data, titles_from_data=True)
        bar.set_categories(labels)
        bar.title = "十大客戶收入"
        bar.y_axis.title = "收入 HKD"
        bar.height = 8
        bar.width = 16
        dashboard.add_chart(bar, "D20")

    if top_product_sheet.max_row >= 2:
        bar = BarChart()
        data = Reference(top_product_sheet, min_col=4, min_row=1, max_row=top_product_sheet.max_row)
        labels = Reference(top_product_sheet, min_col=2, min_row=2, max_row=top_product_sheet.max_row)
        bar.add_data(data, titles_from_data=True)
        bar.set_categories(labels)
        bar.title = "十大貨品收入"
        bar.y_axis.title = "收入 HKD"
        bar.height = 8
        bar.width = 16
        dashboard.add_chart(bar, "N20")

    if comparison_sheet is not None and comparison_sheet.max_row >= 2:
        comparison = BarChart()
        data = Reference(comparison_sheet, min_col=2, max_col=3, min_row=1, max_row=comparison_sheet.max_row)
        labels = Reference(comparison_sheet, min_col=1, min_row=2, max_row=comparison_sheet.max_row)
        comparison.add_data(data, titles_from_data=True)
        comparison.set_categories(labels)
        comparison.title = "本月與上月比較"
        comparison.height = 9
        comparison.width = 18
        dashboard.add_chart(comparison, "D38")

    autosize(dashboard)


def write_report(output_path: Path, sale_rows: list[SaleRow], issues: list[ImportIssue]):
    workbook = Workbook()
    workbook.remove(workbook.active)

    total_revenue = sum((row.revenue for row in sale_rows), Decimal("0"))
    total_quantity = sum((row.quantity for row in sale_rows), Decimal("0"))
    dates = sorted({row.order_date for row in sale_rows if row.order_date})
    report_date = dates[0].isoformat() if len(dates) == 1 else "Multiple / unknown"

    summary = [
        ["報表日期", report_date],
        ["總訂單數量", len({row.source_file for row in sale_rows})],
        ["售出貨品記錄數", len(sale_rows)],
        ["總售出数量", float(total_quantity)],
        ["總收入 HKD", float(total_revenue)],
        ["問題數量", len(issues)],
    ]
    add_sheet(workbook, "總覽", ["項目", "數值"], summary)

    grouped = defaultdict(lambda: {"quantity": Decimal("0"), "revenue": Decimal("0"), "rows": 0})
    for row in sale_rows:
        key = row.customer
        grouped[key]["quantity"] += row.quantity
        grouped[key]["revenue"] += row.revenue
        grouped[key]["rows"] += 1
    add_sheet(
        workbook,
        "按客戶",
        ["客戶", "售出貨品記錄數", "售出数量", "收入 HKD"],
        [[key, val["rows"], float(val["quantity"]), float(val["revenue"])] for key, val in sorted(grouped.items())],
    )

    grouped.clear()
    for row in sale_rows:
        key = (row.category, row.product)
        grouped[key]["quantity"] += row.quantity
        grouped[key]["revenue"] += row.revenue
        grouped[key]["rows"] += 1
    add_sheet(
        workbook,
        "按貨品",
        ["分類", "貨品", "售出貨品記錄數", "售出数量", "收入 HKD"],
        [[cat, product, val["rows"], float(val["quantity"]), float(val["revenue"])] for (cat, product), val in sorted(grouped.items())],
    )

    grouped.clear()
    for row in sale_rows:
        grouped[row.category]["quantity"] += row.quantity
        grouped[row.category]["revenue"] += row.revenue
        grouped[row.category]["rows"] += 1
    add_sheet(
        workbook,
        "按分類",
        ["分類", "售出貨品記錄數", "售出数量", "收入 HKD"],
        [[key, val["rows"], float(val["quantity"]), float(val["revenue"])] for key, val in sorted(grouped.items())],
    )

    add_sheet(
        workbook,
        "明細",
        [
            "來源檔案",
            "日期",
            "客戶",
            "發貨單編號",
            "分類",
            "貨品",
            "部位",
            "售出数量",
            "售價",
            "售價類型",
            "平均買入價",
            "收入",
        ],
        [
            [
                row.source_file,
                row.order_date.isoformat() if row.order_date else "",
                row.customer,
                row.order_number,
                row.category,
                row.product,
                row.part,
                float(row.quantity),
                float(row.unit_price),
                row.price_type,
                float(row.buy_in_price) if row.buy_in_price is not None else "",
                float(row.revenue),
            ]
            for row in sale_rows
        ],
    )

    add_sheet(
        workbook,
        "問題",
        ["來源檔案", "嚴重程度", "問題"],
        [[issue.source_file, issue.severity, issue.message] for issue in issues],
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def percent_change(current: Decimal, previous: Decimal):
    if previous == 0:
        return "沒有上月資料" if current != 0 else 0
    return float(money((current - previous) / previous * Decimal("100")) or Decimal("0"))


def write_monthly_report(
    output_path: Path,
    month: str,
    sale_rows: list[SaleRow],
    issues: list[ImportIssue],
    previous_rows: list[SaleRow] | None = None,
):
    previous_rows = previous_rows or []
    workbook = Workbook()
    workbook.remove(workbook.active)

    total_revenue = sum((row.revenue for row in sale_rows), Decimal("0"))
    total_quantity = sum((row.quantity for row in sale_rows), Decimal("0"))
    order_files = len({row.source_file for row in sale_rows})
    customers = len({row.customer for row in sale_rows if row.customer})
    avg_order = money(total_revenue / Decimal(order_files)) if order_files else Decimal("0")
    avg_price = money(total_revenue / total_quantity) if total_quantity else Decimal("0")

    summary = [
        ["報表月份", month],
        ["總訂單數量", order_files],
        ["客戶數量", customers],
        ["售出貨品記錄數", len(sale_rows)],
        ["總售出数量", float(total_quantity)],
        ["總收入 HKD", float(total_revenue)],
        ["平均每單收入 HKD", float(avg_order or Decimal("0"))],
        ["平均售價 HKD", float(avg_price or Decimal("0"))],
        ["問題數量", len(issues)],
    ]
    summary_sheet = add_sheet(workbook, "總覽", ["項目", "數值"], summary)

    grouped = defaultdict(lambda: {"quantity": Decimal("0"), "revenue": Decimal("0"), "rows": 0, "files": set()})
    for row in sale_rows:
        grouped[row.customer]["quantity"] += row.quantity
        grouped[row.customer]["revenue"] += row.revenue
        grouped[row.customer]["rows"] += 1
        grouped[row.customer]["files"].add(row.source_file)
    customer_sheet = add_sheet(
        workbook,
        "客戶月報",
        ["客戶", "總訂單數量", "售出貨品記錄數", "售出数量", "收入 HKD", "平均售價 HKD", "收入佔比 %"],
        [
            [
                key,
                len(val["files"]),
                val["rows"],
                float(val["quantity"]),
                float(val["revenue"]),
                float(money(val["revenue"] / val["quantity"]) or Decimal("0")) if val["quantity"] else "",
                float(money(val["revenue"] / total_revenue * Decimal("100")) or Decimal("0")) if total_revenue else 0,
            ]
            for key, val in sorted(grouped.items(), key=lambda item: item[1]["revenue"], reverse=True)
        ],
    )

    grouped.clear()
    for row in sale_rows:
        key = (row.category, row.product)
        grouped[key]["quantity"] += row.quantity
        grouped[key]["revenue"] += row.revenue
        grouped[key]["rows"] += 1
    product_sheet = add_sheet(
        workbook,
        "貨品月報",
        ["分類", "貨品", "售出貨品記錄數", "售出数量", "收入 HKD", "平均售價 HKD", "收入佔比 %"],
        [
            [
                cat,
                product,
                val["rows"],
                float(val["quantity"]),
                float(val["revenue"]),
                float(money(val["revenue"] / val["quantity"]) or Decimal("0")) if val["quantity"] else "",
                float(money(val["revenue"] / total_revenue * Decimal("100")) or Decimal("0")) if total_revenue else 0,
            ]
            for (cat, product), val in sorted(grouped.items(), key=lambda item: item[1]["revenue"], reverse=True)
        ],
    )

    grouped.clear()
    for row in sale_rows:
        grouped[row.category]["quantity"] += row.quantity
        grouped[row.category]["revenue"] += row.revenue
        grouped[row.category]["rows"] += 1
    category_sheet = add_sheet(
        workbook,
        "分類月報",
        ["分類", "售出貨品記錄數", "售出数量", "收入 HKD", "收入佔比 %"],
        [
            [
                key,
                val["rows"],
                float(val["quantity"]),
                float(val["revenue"]),
                float(money(val["revenue"] / total_revenue * Decimal("100")) or Decimal("0")) if total_revenue else 0,
            ]
            for key, val in sorted(grouped.items(), key=lambda item: item[1]["revenue"], reverse=True)
        ],
    )

    grouped.clear()
    for row in sale_rows:
        key = row.order_date.isoformat() if row.order_date else ""
        grouped[key]["quantity"] += row.quantity
        grouped[key]["revenue"] += row.revenue
        grouped[key]["rows"] += 1
        grouped[key]["files"].add(row.source_file)
    trend_sheet = add_sheet(
        workbook,
        "每日走勢",
        ["日期", "總訂單數量", "售出貨品記錄數", "售出数量", "收入 HKD"],
        [
            [key, len(val["files"]), val["rows"], float(val["quantity"]), float(val["revenue"])]
            for key, val in sorted(grouped.items())
        ],
    )

    customer_revenue = defaultdict(lambda: Decimal("0"))
    for row in sale_rows:
        customer_revenue[row.customer] += row.revenue
    top_customer_sheet = add_sheet(
        workbook,
        "十大客戶",
        ["客戶", "收入 HKD", "收入佔比 %"],
        [
            [
                customer,
                float(revenue),
                float(money(revenue / total_revenue * Decimal("100")) or Decimal("0")) if total_revenue else 0,
            ]
            for customer, revenue in sorted(customer_revenue.items(), key=lambda item: item[1], reverse=True)[:10]
        ],
    )

    product_revenue = defaultdict(lambda: {"quantity": Decimal("0"), "revenue": Decimal("0")})
    for row in sale_rows:
        key = (row.category, row.product)
        product_revenue[key]["quantity"] += row.quantity
        product_revenue[key]["revenue"] += row.revenue
    top_product_sheet = add_sheet(
        workbook,
        "十大貨品",
        ["分類", "貨品", "售出数量", "收入 HKD", "平均售價 HKD"],
        [
            [
                category,
                product,
                float(val["quantity"]),
                float(val["revenue"]),
                float(money(val["revenue"] / val["quantity"]) or Decimal("0")) if val["quantity"] else "",
            ]
            for (category, product), val in sorted(product_revenue.items(), key=lambda item: item[1]["revenue"], reverse=True)[:10]
        ],
    )

    issue_group = defaultdict(int)
    for issue in issues:
        issue_group[issue.severity] += 1
    issue_summary_sheet = add_sheet(
        workbook,
        "問題摘要",
        ["嚴重程度", "數量"],
        [[severity, count] for severity, count in sorted(issue_group.items())],
    )

    previous_revenue = sum((row.revenue for row in previous_rows), Decimal("0"))
    previous_quantity = sum((row.quantity for row in previous_rows), Decimal("0"))
    previous_orders = len({row.source_file for row in previous_rows})
    previous_customers = len({row.customer for row in previous_rows if row.customer})
    current_avg_order = total_revenue / Decimal(order_files) if order_files else Decimal("0")
    previous_avg_order = previous_revenue / Decimal(previous_orders) if previous_orders else Decimal("0")

    comparison_sheet = add_sheet(
        workbook,
        "月度比較",
        ["指標", "本月", "上月", "變化 %"],
        [
            ["總收入 HKD", float(total_revenue), float(previous_revenue), percent_change(total_revenue, previous_revenue)],
            ["總售出数量", float(total_quantity), float(previous_quantity), percent_change(total_quantity, previous_quantity)],
            ["總訂單數量", order_files, previous_orders, percent_change(Decimal(order_files), Decimal(previous_orders))],
            ["客戶數量", customers, previous_customers, percent_change(Decimal(customers), Decimal(previous_customers))],
            ["平均每單收入 HKD", float(money(current_avg_order) or Decimal("0")), float(money(previous_avg_order) or Decimal("0")), percent_change(current_avg_order, previous_avg_order)],
        ],
    )

    def grouped_values(rows: list[SaleRow], key_fn):
        result = defaultdict(lambda: {"quantity": Decimal("0"), "revenue": Decimal("0")})
        for row in rows:
            key = key_fn(row)
            result[key]["quantity"] += row.quantity
            result[key]["revenue"] += row.revenue
        return result

    current_customer = grouped_values(sale_rows, lambda row: row.customer)
    previous_customer = grouped_values(previous_rows, lambda row: row.customer)
    customer_keys = sorted(set(current_customer) | set(previous_customer))
    add_sheet(
        workbook,
        "客戶比較",
        ["客戶", "本月收入", "上月收入", "收入變化 %", "本月数量", "上月数量", "数量變化 %"],
        [[key, float(current_customer[key]["revenue"]), float(previous_customer[key]["revenue"]), percent_change(current_customer[key]["revenue"], previous_customer[key]["revenue"]), float(current_customer[key]["quantity"]), float(previous_customer[key]["quantity"]), percent_change(current_customer[key]["quantity"], previous_customer[key]["quantity"])] for key in customer_keys],
    )

    current_product = grouped_values(sale_rows, lambda row: (row.category, row.product))
    previous_product = grouped_values(previous_rows, lambda row: (row.category, row.product))
    product_keys = sorted(set(current_product) | set(previous_product))
    add_sheet(
        workbook,
        "貨品比較",
        ["分類", "貨品", "本月收入", "上月收入", "收入變化 %", "本月数量", "上月数量", "数量變化 %"],
        [[key[0], key[1], float(current_product[key]["revenue"]), float(previous_product[key]["revenue"]), percent_change(current_product[key]["revenue"], previous_product[key]["revenue"]), float(current_product[key]["quantity"]), float(previous_product[key]["quantity"]), percent_change(current_product[key]["quantity"], previous_product[key]["quantity"])] for key in product_keys],
    )

    current_category = grouped_values(sale_rows, lambda row: row.category)
    previous_category = grouped_values(previous_rows, lambda row: row.category)
    category_keys = sorted(set(current_category) | set(previous_category))
    add_sheet(
        workbook,
        "分類比較",
        ["分類", "本月收入", "上月收入", "收入變化 %", "本月数量", "上月数量", "数量變化 %"],
        [[key, float(current_category[key]["revenue"]), float(previous_category[key]["revenue"]), percent_change(current_category[key]["revenue"], previous_category[key]["revenue"]), float(current_category[key]["quantity"]), float(previous_category[key]["quantity"]), percent_change(current_category[key]["quantity"], previous_category[key]["quantity"])] for key in category_keys],
    )

    add_monthly_charts(
        workbook,
        summary_sheet,
        category_sheet,
        trend_sheet,
        top_customer_sheet,
        top_product_sheet,
        comparison_sheet,
    )

    add_sheet(
        workbook,
        "明細",
        [
            "來源檔案",
            "日期",
            "客戶",
            "發貨單編號",
            "分類",
            "貨品",
            "部位",
            "售出数量",
            "售價",
            "售價類型",
            "平均買入價",
            "收入",
        ],
        [
            [
                row.source_file,
                row.order_date.isoformat() if row.order_date else "",
                row.customer,
                row.order_number,
                row.category,
                row.product,
                row.part,
                float(row.quantity),
                float(row.unit_price),
                row.price_type,
                float(row.buy_in_price) if row.buy_in_price is not None else "",
                float(row.revenue),
            ]
            for row in sale_rows
        ],
    )

    add_sheet(
        workbook,
        "問題",
        ["來源檔案", "嚴重程度", "問題"],
        [[issue.source_file, issue.severity, issue.message] for issue in issues],
    )

    consolidate_monthly_sheets(workbook)
    reorder_monthly_sheets(workbook)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def reorder_monthly_sheets(workbook: Workbook) -> None:
    preferred_order = [
        "總覽",
        "月度比較",
        "客戶分析",
        "貨品分析",
        "分類分析",
        "每日走勢",
        "明細",
        "問題",
        "客戶比較",
        "貨品比較",
        "分類比較",
        "十大客戶",
        "十大貨品",
        "問題摘要",
    ]
    workbook._sheets = [workbook[name] for name in preferred_order if name in workbook.sheetnames]


def generate_daily_report(folder: Path, output_path: Path) -> tuple[int, int, int]:
    """Read a folder of order workbooks, write report, return counts.

    Returns:
        (file_count, sold_row_count, issue_count)
    """
    all_rows: list[SaleRow] = []
    all_issues: list[ImportIssue] = []
    workbook_paths = list(iter_workbooks(folder))
    expected_folder_date = folder_date(folder)

    for workbook_path in workbook_paths:
        rows, issues = read_order_file(workbook_path, expected_folder_date)
        all_rows.extend(rows)
        all_issues.extend(issues)

    all_issues.extend(detect_duplicate_orders(all_rows))
    write_report(output_path, all_rows, all_issues)
    from report_pdfs import write_daily_pdf

    write_daily_pdf(output_path.with_suffix(".pdf"), all_rows, all_issues)
    return len(workbook_paths), len(all_rows), len(all_issues)


def generate_monthly_report(orders_base_folder: Path, month: str, output_path: Path) -> tuple[int, int, int]:
    """Read all YYYYMMDD folders for a month and generate a monthly report.

    Args:
        orders_base_folder: folder containing day folders like 20260618.
        month: YYYYMM.
        output_path: report path.
    """
    if not re.fullmatch(r"\d{6}", month):
        raise ValueError("month must be YYYYMM")

    all_rows: list[SaleRow] = []
    all_issues: list[ImportIssue] = []
    previous_rows: list[SaleRow] = []
    file_count = 0
    current_month_date = datetime.strptime(month + "01", "%Y%m%d").date()
    previous_month = (current_month_date.replace(day=1) - timedelta(days=1)).strftime("%Y%m")

    for day_folder in sorted(orders_base_folder.iterdir()):
        if not day_folder.is_dir():
            continue
        if not re.fullmatch(r"\d{8}", day_folder.name):
            continue
        folder_month = day_folder.name[:6]
        if folder_month not in {month, previous_month}:
            continue

        expected_folder_date = folder_date(day_folder)
        workbook_paths = list(iter_workbooks(day_folder))
        if folder_month == month:
            file_count += len(workbook_paths)
        for workbook_path in workbook_paths:
            rows, issues = read_order_file(workbook_path, expected_folder_date)
            if folder_month == month:
                all_rows.extend(rows)
                all_issues.extend(issues)
            else:
                previous_rows.extend(rows)

    all_issues.extend(detect_duplicate_orders(all_rows))
    write_monthly_report(output_path, month, all_rows, all_issues, previous_rows)
    from report_pdfs import write_monthly_pdf

    write_monthly_pdf(output_path.with_suffix(".pdf"), month, all_rows, previous_rows, all_issues)
    return file_count, len(all_rows), len(all_issues)


def main() -> int:
    parser = argparse.ArgumentParser(description="Read order Excel files and generate a daily report.")
    parser.add_argument("folder", type=Path, help="Folder containing YYYYMMDD_客戶名_序號.xlsx files")
    parser.add_argument("--output", type=Path, default=None, help="Output .xlsx path")
    args = parser.parse_args()

    if not args.folder.exists() or not args.folder.is_dir():
        raise SystemExit(f"Folder does not exist: {args.folder}")

    if args.output:
        output_path = args.output
    else:
        output_path = args.folder / "daily_report.xlsx"

    file_count, sold_row_count, issue_count = generate_daily_report(args.folder, output_path)

    print(f"Read files: {file_count}")
    print(f"Sold rows: {sold_row_count}")
    print(f"Issues: {issue_count}")
    print(f"Report: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
