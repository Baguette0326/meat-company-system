from __future__ import annotations

import calendar
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


def pdf_pct_change(current: Decimal, previous: Decimal) -> Decimal | None:
    if previous == 0:
        return None
    return (current - previous) / previous * Decimal("100")


def pdf_pct(value: Decimal | None) -> str:
    if value is None:
        return "沒有上月資料"
    return f"{float(value):+.1f}%"


def add_pdf_alert(alerts: list[list[str]], severity: str, title: str, detail: str, action: str) -> None:
    alerts.append([severity, title, detail, action])


def adjustment_counts(rows, issues) -> dict[str, int]:
    return {
        "取消": len({issue.source_file for issue in issues if "訂單狀態為取消" in issue.message}),
        "更正": len({row.source_file for row in rows if row.order_status == "更正"}),
        "退貨": len({row.source_file for row in rows if row.order_status == "退貨"}),
    }


def monthly_alerts(rows, previous_rows, issues) -> list[list[str]]:
    alerts: list[list[str]] = []
    total_revenue = sum((row.revenue for row in rows), Decimal("0"))
    previous_revenue = sum((row.revenue for row in previous_rows), Decimal("0"))
    total_quantity = sum((row.quantity for row in rows), Decimal("0"))
    previous_quantity = sum((row.quantity for row in previous_rows), Decimal("0"))
    current_customers = grouped(rows, lambda row: row.customer)
    previous_customers = grouped(previous_rows, lambda row: row.customer)
    current_products = grouped(rows, lambda row: (row.category, row.product))
    previous_products = grouped(previous_rows, lambda row: (row.category, row.product))
    adjustments = adjustment_counts(rows, issues)

    if not rows:
        add_pdf_alert(alerts, "高", "本月沒有有效銷售資料", "沒有讀到可計算的售出貨品。", "檢查訂單資料夾。")
    if issues:
        add_pdf_alert(alerts, "中", "有輸入問題需要檢查", f"共有 {len(issues)} 個提示/警告/錯誤。", "查看 Excel「問題」分頁。")
    if any(adjustments.values()):
        add_pdf_alert(alerts, "中", "本月有訂單調整", f"取消 {adjustments['取消']}，更正 {adjustments['更正']}，退貨 {adjustments['退貨']}。", "確認調整原因。")

    revenue_change = pdf_pct_change(total_revenue, previous_revenue)
    if revenue_change is not None:
        if revenue_change <= Decimal("-20"):
            add_pdf_alert(alerts, "高", "收入明顯下降", f"較上月 {pdf_pct(revenue_change)}。", "檢查主要客戶及貨品。")
        elif revenue_change >= Decimal("20"):
            add_pdf_alert(alerts, "低", "收入明顯上升", f"較上月 {pdf_pct(revenue_change)}。", "確認增長來源。")

    quantity_change = pdf_pct_change(total_quantity, previous_quantity)
    if quantity_change is not None:
        if quantity_change <= Decimal("-20"):
            add_pdf_alert(alerts, "高", "銷量明顯下降", f"較上月 {pdf_pct(quantity_change)}。", "檢查需求或缺貨。")
        elif quantity_change >= Decimal("20"):
            add_pdf_alert(alerts, "低", "銷量明顯上升", f"較上月 {pdf_pct(quantity_change)}。", "留意庫存供應。")

    lost_customers = [
        customer
        for customer, previous in previous_customers.items()
        if previous["revenue"] > 0 and current_customers[customer]["revenue"] == 0
    ]
    if lost_customers:
        add_pdf_alert(alerts, "高", "有客戶本月沒有再下單", f"{len(lost_customers)} 位上月客戶本月沒有收入。", "跟進客戶狀況。")

    changed_products = []
    for product, previous in previous_products.items():
        if previous["quantity"] == 0:
            continue
        ratio = pdf_pct_change(current_products[product]["quantity"], previous["quantity"])
        if ratio is not None and abs(ratio) >= Decimal("30"):
            changed_products.append((product, ratio))
    if changed_products:
        product, ratio = sorted(changed_products, key=lambda item: abs(item[1]), reverse=True)[0]
        add_pdf_alert(alerts, "中", "貨品需求變化明顯", f"{product[1]} 售出数量較上月 {pdf_pct(ratio)}。", "查看貨品分析。")

    if total_revenue > 0 and current_customers:
        top_customer, top_value = max(current_customers.items(), key=lambda item: item[1]["revenue"])
        share = top_value["revenue"] / total_revenue * Decimal("100")
        if share >= Decimal("50"):
            add_pdf_alert(alerts, "中", "收入集中在單一客戶", f"{top_customer} 佔本月收入 {float(share):.1f}%。", "留意客戶集中風險。")

    if not alerts:
        add_pdf_alert(alerts, "正常", "沒有明顯異常", "本月沒有觸發管理提示。", "可正常查看月報。")
    return alerts[:8]


def confidence_label(elapsed_days: int, order_count: int, previous_rows) -> str:
    if elapsed_days >= 14 and order_count >= 20 and previous_rows:
        return "高"
    if elapsed_days >= 7 and order_count >= 5:
        return "中"
    return "低"


def forecast_next(current_month_end: Decimal, previous_value: Decimal) -> Decimal:
    if previous_value <= 0:
        return current_month_end
    trend = (current_month_end - previous_value) / previous_value
    trend = max(Decimal("-0.30"), min(Decimal("0.30"), trend))
    return current_month_end * (Decimal("1") + trend / Decimal("2"))


def monthly_forecast_rows(month: str, rows, previous_rows) -> list[list[str]]:
    year = int(month[:4])
    month_number = int(month[4:])
    days_in_month = calendar.monthrange(year, month_number)[1]
    dates = [row.order_date for row in rows if row.order_date]
    elapsed_days = max((day.day for day in dates), default=0)
    elapsed_days = max(1, min(elapsed_days, days_in_month))
    scale = Decimal(days_in_month) / Decimal(elapsed_days)
    order_count = len({row.source_file for row in rows})
    previous_order_count = len({row.source_file for row in previous_rows})
    confidence = confidence_label(elapsed_days, order_count, previous_rows)

    total_revenue = sum((row.revenue for row in rows), Decimal("0"))
    total_quantity = sum((row.quantity for row in rows), Decimal("0"))
    previous_revenue = sum((row.revenue for row in previous_rows), Decimal("0"))
    previous_quantity = sum((row.quantity for row in previous_rows), Decimal("0"))

    forecast_revenue = total_revenue * scale
    forecast_quantity = total_quantity * scale
    forecast_orders = Decimal(order_count) * scale

    result = [
        ["總收入", money(total_revenue), money(forecast_revenue), money(forecast_next(forecast_revenue, previous_revenue)), confidence],
        ["總售出数量", quantity(total_quantity), quantity(forecast_quantity), quantity(forecast_next(forecast_quantity, previous_quantity)), confidence],
        ["總訂單數量", str(order_count), f"{float(forecast_orders):.1f}", f"{float(forecast_next(forecast_orders, Decimal(previous_order_count))):.1f}", confidence],
    ]

    products = grouped(rows, lambda row: (row.category, row.product))
    previous_products = grouped(previous_rows, lambda row: (row.category, row.product))
    for (category, product), current in sorted(products.items(), key=lambda item: item[1]["revenue"], reverse=True)[:5]:
        previous = previous_products[(category, product)]
        month_end_quantity = current["quantity"] * scale
        next_quantity = forecast_next(month_end_quantity, previous["quantity"])
        result.append([f"{category}-{product}", quantity(current["quantity"]), quantity(month_end_quantity), quantity(next_quantity), confidence])
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

    draw_title(pdf, "管理提示", month)
    alert_table = make_table(
        "老闆重點提示",
        ["程度", "提示", "內容", "建議"],
        monthly_alerts(rows, previous_rows, issues),
        [20 * mm, 46 * mm, 96 * mm, 96 * mm],
    )
    current_y = PAGE[1] - 42 * mm
    for flowable in alert_table:
        _, height = flowable.wrap(PAGE[0] - 28 * mm, PAGE[1])
        flowable.drawOn(pdf, 14 * mm, current_y - height)
        current_y -= height + 3 * mm
    pdf.setFillColor(GRAY)
    pdf.setFont(FONT_NAME, 9)
    pdf.drawString(14 * mm, 10 * mm, "提示只作管理參考，詳細資料以 Excel 月報各分頁為準。")
    pdf.showPage()

    draw_title(pdf, "銷售預測", month)
    forecast_table = make_table(
        "月底及下月預測",
        ["項目", "目前資料", "本月月底預測", "下月預測", "信心"],
        monthly_forecast_rows(month, rows, previous_rows),
        [58 * mm, 42 * mm, 52 * mm, 52 * mm, 24 * mm],
    )
    current_y = PAGE[1] - 42 * mm
    for flowable in forecast_table:
        _, height = flowable.wrap(PAGE[0] - 28 * mm, PAGE[1])
        flowable.drawOn(pdf, 14 * mm, current_y - height)
        current_y -= height + 3 * mm
    pdf.setFillColor(GRAY)
    pdf.setFont(FONT_NAME, 9)
    pdf.drawString(14 * mm, 10 * mm, "預測根據現有訂單速度及上月趨勢估算；新系統初期資料少，請只作管理參考。")
    pdf.showPage()
    pdf.save()
