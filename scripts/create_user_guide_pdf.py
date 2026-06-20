from __future__ import annotations

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "銷售報表系統_使用指南.pdf"
PAGE = landscape(A4)
FONT_NAME = "MicrosoftJhengHei"
FONT_PATH = r"C:\Windows\Fonts\msjh.ttc"


def draw_paragraph(pdf, text: str, style: ParagraphStyle, x: float, y: float, width: float) -> float:
    paragraph = Paragraph(text, style)
    _, height = paragraph.wrap(width, PAGE[1])
    paragraph.drawOn(pdf, x, y - height)
    return y - height


def draw_section(pdf, title: str, body: str, x: float, y: float, width: float, styles) -> float:
    y = draw_paragraph(pdf, title, styles["section"], x, y, width)
    y -= 2 * mm
    y = draw_paragraph(pdf, body, styles["body"], x, y, width)
    return y - 4 * mm


def main() -> None:
    pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH, subfontIndex=0))
    pdf = canvas.Canvas(str(OUTPUT), pagesize=PAGE)

    styles = {
        "title": ParagraphStyle(
            "title",
            fontName=FONT_NAME,
            fontSize=20,
            leading=24,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#115E59"),
        ),
        "section": ParagraphStyle(
            "section",
            fontName=FONT_NAME,
            fontSize=12,
            leading=16,
            textColor=colors.white,
            backColor=colors.HexColor("#0F766E"),
            borderPadding=(4, 6, 4, 6),
        ),
        "body": ParagraphStyle(
            "body",
            fontName=FONT_NAME,
            fontSize=9.3,
            leading=13.2,
            alignment=TA_LEFT,
            textColor=colors.HexColor("#1F2937"),
        ),
        "note": ParagraphStyle(
            "note",
            fontName=FONT_NAME,
            fontSize=9.3,
            leading=13.2,
            textColor=colors.HexColor("#92400E"),
            backColor=colors.HexColor("#FEF3C7"),
            borderPadding=6,
        ),
    }

    page_width, page_height = PAGE
    margin = 12 * mm
    title_y = page_height - 9 * mm
    draw_paragraph(pdf, "銷售報表系統｜一頁使用指南", styles["title"], margin, title_y, page_width - 2 * margin)

    gap = 8 * mm
    col_width = (page_width - 2 * margin - 2 * gap) / 3
    xs = [margin, margin + col_width + gap, margin + 2 * (col_width + gap)]
    top = page_height - 24 * mm

    y = top
    y = draw_section(pdf, "每日操作", "1. 按「打開今日訂單資料夾」<br/>2. 把當日訂單 Excel 放入資料夾<br/>3. 按「產生今日報表」<br/>4. 按「打開今日報表」", xs[0], y, col_width, styles)
    y = draw_section(pdf, "月報操作", "1. 按「產生本月月報」<br/>2. 按「打開本月月報」<br/><br/>月報會自動整合 Orders 內該月份所有日期資料夾。", xs[0], y, col_width, styles)
    y = draw_section(pdf, "檔案命名", "訂單：YYYYMMDD_客戶名_序號.xlsx<br/>例：20260619_海景酒家_001.xlsx<br/><br/>每日報表：每日報表_YYYYMMDD.xlsx<br/>月報表：月報表_YYYYMM.xlsx", xs[0], y, col_width, styles)

    y = top
    y = draw_section(pdf, "每日報表內容", "• 總訂單數量<br/>• 總售出数量及總收入<br/>• 按客戶、貨品及分類統計<br/>• 銷售明細<br/>• 輸入問題及重複訂單警告", xs[1], y, col_width, styles)
    y = draw_section(pdf, "月報表內容", "• 本月與上月比較<br/>• 客戶、貨品及分類分析<br/>• 十大客戶及十大貨品<br/>• 每日收入走勢<br/>• 平均售價及收入佔比<br/>• 自動 Excel 表格及圖表", xs[1], y, col_width, styles)
    y = draw_section(pdf, "月報主要頁面", "總覽｜月度比較｜客戶分析｜貨品分析<br/>分類分析｜每日走勢｜明細｜問題", xs[1], y, col_width, styles)

    y = top
    y = draw_section(pdf, "什麼是「售出貨品記錄數」？", "代表有多少行貨品銷售記錄。<br/><br/>例：一張訂單有牛柳、牛腩、豬扒三個貨品，便有 3 個貨品記錄。<br/><br/>它不是總訂單數，也不是售出 KG 数量。", xs[2], y, col_width, styles)
    y = draw_section(pdf, "什麼是「總訂單數量」？", "代表系統成功讀取多少張訂單 Excel。<br/><br/>每張訂單 Excel 代表一個客戶訂單。<br/><br/>例：當日資料夾有 12 張有效訂單 Excel，總訂單數量便是 12。", xs[2], y, col_width, styles)
    draw_paragraph(pdf, "重要：更新報表前先關閉已開啟的報表 Excel。若報表顯示問題，請到「問題」頁查看來源檔案及原因。", styles["note"], xs[2], y, col_width)

    pdf.showPage()
    pdf.save()
    print(f"created={OUTPUT}")


if __name__ == "__main__":
    main()
