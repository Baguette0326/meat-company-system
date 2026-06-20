from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Table, TableStyle


PAGE = landscape(A4)
FONT_NAME = "MicrosoftJhengHei"
FONT_PATH = r"C:\Windows\Fonts\msjh.ttc"
TEAL = colors.HexColor("#0F766E")
LIGHT = colors.HexColor("#ECFDF5")
GRAY = colors.HexColor("#475569")


def setup_fonts() -> None:
    try:
        pdfmetrics.getFont(FONT_NAME)
    except KeyError:
        pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH, subfontIndex=0))


def money(value: Decimal) -> str:
    return f"HKD {float(value):,.2f}"


def quantity(value: Decimal) -> str:
    return f"{float(value):,.2f}"


def draw_title(pdf, title: str, subtitle: str) -> None:
    pdf.setFillColor(TEAL)
    pdf.rect(0, PAGE[1] - 27 * mm, PAGE[0], 27 * mm, fill=1, stroke=0)
    pdf.setFillColor(colors.white)
    pdf.setFont(FONT_NAME, 21)
    pdf.drawString(14 * mm, PAGE[1] - 15 * mm, title)
    pdf.setFont(FONT_NAME, 11)
    pdf.drawRightString(PAGE[0] - 14 * mm, PAGE[1] - 15 * mm, subtitle)


def draw_kpis(pdf, values: list[tuple[str, str]], y: float) -> float:
    margin = 14 * mm
    gap = 5 * mm
    width = (PAGE[0] - 2 * margin - gap * (len(values) - 1)) / len(values)
    for index, (label, value) in enumerate(values):
        x = margin + index * (width + gap)
        pdf.setFillColor(LIGHT)
        pdf.roundRect(x, y - 22 * mm, width, 22 * mm, 3 * mm, fill=1, stroke=0)
        pdf.setFillColor(GRAY)
        pdf.setFont(FONT_NAME, 9)
        pdf.drawString(x + 5 * mm, y - 7 * mm, label)
        pdf.setFillColor(TEAL)
        pdf.setFont(FONT_NAME, 16)
        pdf.drawString(x + 5 * mm, y - 17 * mm, value)
    return y - 28 * mm


def make_table(title: str, headers: list[str], data: list[list], widths: list[float]) -> list:
    title_style = ParagraphStyle("table_title", fontName=FONT_NAME, fontSize=12, leading=15, textColor=TEAL)
    body_style = ParagraphStyle("table_body", fontName=FONT_NAME, fontSize=8.5, leading=11, textColor=colors.HexColor("#1F2937"))
    rows = [[Paragraph(str(value), body_style) for value in headers]]
    rows.extend([[Paragraph(str(value), body_style) for value in row] for row in data])
    table = Table(rows, colWidths=widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), TEAL),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#CBD5E1")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return [Paragraph(title, title_style), table]


def draw_three_tables(pdf, tables: list[list], y: float) -> float:
    margin = 14 * mm
    gap = 6 * mm
    col_width = (PAGE[0] - 2 * margin - 2 * gap) / 3
    for index, flowables in enumerate(tables):
        x = margin + index * (col_width + gap)
        current_y = y
        for flowable in flowables:
            _, height = flowable.wrap(col_width, PAGE[1])
            flowable.drawOn(pdf, x, current_y - height)
            current_y -= height + 2 * mm
    return y


def grouped(rows, key_fn):
    result = defaultdict(lambda: {"quantity": Decimal("0"), "revenue": Decimal("0"), "files": set()})
    for row in rows:
        key = key_fn(row)
        result[key]["quantity"] += row.quantity
        result[key]["revenue"] += row.revenue
        result[key]["files"].add(row.source_file)
    return result


def write_daily_pdf(output: Path, rows, issues) -> None:
    setup_fonts()
    pdf = canvas.Canvas(str(output), pagesize=PAGE)
    dates = sorted({row.order_date for row in rows if row.order_date})
    report_date = dates[0].isoformat() if len(dates) == 1 else "日期不明"
    total_quantity = sum((row.quantity for row in rows), Decimal("0"))
    total_revenue = sum((row.revenue for row in rows), Decimal("0"))
    order_count = len({row.source_file for row in rows})

    draw_title(pdf, "每日銷售報表", report_date)
    y = draw_kpis(pdf, [
        ("總訂單數量", str(order_count)),
        ("售出貨品記錄數", str(len(rows))),
        ("總售出數量", quantity(total_quantity)),
        ("總收入", money(total_revenue)),
        ("問題數量", str(len(issues))),
    ], PAGE[1] - 35 * mm)

    customers = grouped(rows, lambda row: row.customer)
    products = grouped(rows, lambda row: row.product)
    categories = grouped(rows, lambda row: row.category)
    customer_data = [[key, len(val["files"]), money(val["revenue"])] for key, val in sorted(customers.items(), key=lambda item: item[1]["revenue"], reverse=True)[:8]]
    product_data = [[key, quantity(val["quantity"]), money(val["revenue"])] for key, val in sorted(products.items(), key=lambda item: item[1]["revenue"], reverse=True)[:8]]
    category_data = [[key, quantity(val["quantity"]), money(val["revenue"])] for key, val in sorted(categories.items(), key=lambda item: item[1]["revenue"], reverse=True)]
    draw_three_tables(pdf, [
        make_table("主要客戶", ["客戶", "訂單", "收入"], customer_data, [40 * mm, 18 * mm, 28 * mm]),
        make_table("主要貨品", ["貨品", "數量", "收入"], product_data, [40 * mm, 20 * mm, 28 * mm]),
        make_table("分類總計", ["分類", "數量", "收入"], category_data, [31 * mm, 25 * mm, 31 * mm]),
    ], y)

    pdf.setFont(FONT_NAME, 9)
    pdf.setFillColor(colors.HexColor("#92400E") if issues else GRAY)
    pdf.drawString(14 * mm, 10 * mm, f"問題數量：{len(issues)}。詳細問題請查看 Excel 報表的「問題」頁。")
    pdf.showPage()
    pdf.save()


def pct_change(current: Decimal, previous: Decimal) -> str:
    if previous == 0:
        return "沒有上月資料"
    return f"{float((current - previous) / previous * Decimal('100')):+.1f}%"


def write_monthly_pdf(output: Path, month: str, rows, previous_rows, issues) -> None:
    setup_fonts()
    pdf = canvas.Canvas(str(output), pagesize=PAGE)
    total_quantity = sum((row.quantity for row in rows), Decimal("0"))
    total_revenue = sum((row.revenue for row in rows), Decimal("0"))
    previous_quantity = sum((row.quantity for row in previous_rows), Decimal("0"))
    previous_revenue = sum((row.revenue for row in previous_rows), Decimal("0"))
    order_count = len({row.source_file for row in rows})

    draw_title(pdf, "月度銷售報表", month)
    y = draw_kpis(pdf, [
        ("總訂單數量", str(order_count)),
        ("客戶數量", str(len({row.customer for row in rows}))),
        ("總售出數量", quantity(total_quantity)),
        ("總收入", money(total_revenue)),
        ("收入較上月", pct_change(total_revenue, previous_revenue)),
    ], PAGE[1] - 35 * mm)

    customers = grouped(rows, lambda row: row.customer)
    products = grouped(rows, lambda row: row.product)
    categories = grouped(rows, lambda row: row.category)
    customer_data = [[key, len(val["files"]), money(val["revenue"])] for key, val in sorted(customers.items(), key=lambda item: item[1]["revenue"], reverse=True)[:8]]
    product_data = [[key, quantity(val["quantity"]), money(val["revenue"])] for key, val in sorted(products.items(), key=lambda item: item[1]["revenue"], reverse=True)[:8]]
    category_data = [[key, quantity(val["quantity"]), money(val["revenue"])] for key, val in sorted(categories.items(), key=lambda item: item[1]["revenue"], reverse=True)]
    draw_three_tables(pdf, [
        make_table("十大客戶（最多 8 項）", ["客戶", "訂單", "收入"], customer_data, [40 * mm, 18 * mm, 28 * mm]),
        make_table("十大貨品（最多 8 項）", ["貨品", "數量", "收入"], product_data, [40 * mm, 20 * mm, 28 * mm]),
        make_table("分類表現", ["分類", "數量", "收入"], category_data, [31 * mm, 25 * mm, 31 * mm]),
    ], y)

    pdf.setFont(FONT_NAME, 9)
    pdf.setFillColor(GRAY)
    pdf.drawString(14 * mm, 10 * mm, f"數量較上月：{pct_change(total_quantity, previous_quantity)}｜問題數量：{len(issues)}。詳細資料請查看 Excel 月報。")
    pdf.showPage()

    previous_order_count = len({row.source_file for row in previous_rows})
    current_customers = len({row.customer for row in rows})
    previous_customers = len({row.customer for row in previous_rows})
    average_order = total_revenue / Decimal(order_count) if order_count else Decimal("0")
    average_price = total_revenue / total_quantity if total_quantity else Decimal("0")

    daily = grouped(rows, lambda row: row.order_date.isoformat() if row.order_date else "日期不明")
    active_days = len(daily)
    repeat_customers = sum(1 for val in customers.values() if len(val["files"]) > 1)
    unique_products = len(products)

    customer_revenues = sorted((val["revenue"] for val in customers.values()), reverse=True)
    product_revenues = sorted((val["revenue"] for val in products.values()), reverse=True)
    top_three_customer_share = sum(customer_revenues[:3], Decimal("0")) / total_revenue * Decimal("100") if total_revenue else Decimal("0")
    top_three_product_share = sum(product_revenues[:3], Decimal("0")) / total_revenue * Decimal("100") if total_revenue else Decimal("0")

    draw_title(pdf, "月度進階統計", month)
    y = draw_kpis(pdf, [
        ("平均每單收入", money(average_order)),
        ("平均售價", money(average_price)),
        ("有銷售日數", str(active_days)),
        ("重複購買客戶", str(repeat_customers)),
        ("售出貨品種類", str(unique_products)),
    ], PAGE[1] - 35 * mm)

    if daily:
        best_revenue_day, best_revenue_val = max(daily.items(), key=lambda item: item[1]["revenue"])
        best_quantity_day, best_quantity_val = max(daily.items(), key=lambda item: item[1]["quantity"])
    else:
        best_revenue_day, best_revenue_val = "沒有資料", {"revenue": Decimal("0")}
        best_quantity_day, best_quantity_val = "沒有資料", {"quantity": Decimal("0")}

    daily_stats = [
        ["最高收入日", best_revenue_day, money(best_revenue_val["revenue"])],
        ["最高銷量日", best_quantity_day, quantity(best_quantity_val["quantity"])],
        ["平均每日收入", "", money(total_revenue / Decimal(active_days) if active_days else Decimal("0"))],
        ["平均每日銷量", "", quantity(total_quantity / Decimal(active_days) if active_days else Decimal("0"))],
        ["平均每日訂單", "", f"{order_count / active_days:.1f}" if active_days else "0"],
    ]

    price_types = grouped(rows, lambda row: row.price_type)
    price_type_data = [
        [key, quantity(val["quantity"]), money(val["revenue"])]
        for key, val in sorted(price_types.items(), key=lambda item: item[1]["revenue"], reverse=True)
    ]

    comparison_data = [
        ["總收入", money(total_revenue), money(previous_revenue), pct_change(total_revenue, previous_revenue)],
        ["總售出數量", quantity(total_quantity), quantity(previous_quantity), pct_change(total_quantity, previous_quantity)],
        ["總訂單數量", str(order_count), str(previous_order_count), pct_change(Decimal(order_count), Decimal(previous_order_count))],
        ["客戶數量", str(current_customers), str(previous_customers), pct_change(Decimal(current_customers), Decimal(previous_customers))],
    ]

    draw_three_tables(pdf, [
        make_table("每日表現", ["項目", "日期", "數值"], daily_stats, [34 * mm, 25 * mm, 29 * mm]),
        make_table("售價類型", ["類型", "數量", "收入"], price_type_data, [30 * mm, 25 * mm, 33 * mm]),
        make_table("本月與上月", ["指標", "本月", "上月", "變化"], comparison_data, [30 * mm, 20 * mm, 20 * mm, 18 * mm]),
    ], y)

    pdf.setFont(FONT_NAME, 10)
    pdf.setFillColor(TEAL)
    pdf.drawString(14 * mm, 25 * mm, f"前三大客戶收入佔比：{float(top_three_customer_share):.1f}%")
    pdf.drawString(105 * mm, 25 * mm, f"前三大貨品收入佔比：{float(top_three_product_share):.1f}%")
    pdf.setFillColor(GRAY)
    pdf.setFont(FONT_NAME, 9)
    pdf.drawString(14 * mm, 10 * mm, f"問題數量：{len(issues)}。完整客戶、貨品、分類、每日走勢及問題資料請查看 Excel 月報。")
    pdf.showPage()
    pdf.save()
