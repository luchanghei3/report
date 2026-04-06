# -*- coding: utf-8 -*-
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Flowable
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import ParagraphStyle
import os

# =========================
# 字体注册
# =========================
FONT_PATH_1 = "C:/Windows/Fonts/simhei.ttf"
FONT_PATH_2 = "C:/Windows/Fonts/simsun.ttc"

if os.path.exists(FONT_PATH_1):
    pdfmetrics.registerFont(TTFont("SimHei", FONT_PATH_1))
if os.path.exists(FONT_PATH_2):
    pdfmetrics.registerFont(TTFont("SimSun", FONT_PATH_2))

FONT = "SimHei"

# =========================
# 表格样式
# =========================
def get_table_style():
    return TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.white),
        ("PADDING", (0, 0), (-1, -1), 6),
        ("BACKGROUND", (0, 0), (-1, 0), colors.Color(0.45, 0.2, 0.6, 1)),  # 表头深紫
        ("BACKGROUND", (0, 1), (-1, -1), colors.Color(0.72, 0.58, 0.85, 1)),  # 内容浅紫
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ])

def get_table_text_style():
    return ParagraphStyle(
        name="TableCell",
        fontName=FONT,
        fontSize=11,
        leading=14,
        alignment=1,
        textColor=colors.white
    )

# =========================
# 菱形标题 Flowable
# =========================
class SectionTitle(Flowable):
    def __init__(self, title, number=None, x=0, y=0,
                 left_color=colors.Color(0.72, 0.58, 0.85, 1),
                 right_color=colors.Color(0.34, 0.28, 0.63, 1),
                 size=30, gap=-5):
        super().__init__()
        self.title = title
        self.number = number
        self.x = x
        self.y = y
        self.size = size
        self.gap = gap
        self.left_color = left_color
        self.right_color = right_color

    def draw(self):
        c = self.canv
        def draw_diamond(px, py, color):
            c.setFillColor(color)
            c.saveState()
            c.translate(px, py)
            c.rotate(45)
            c.rect(-self.size/2, -self.size/2, self.size, self.size, fill=1, stroke=0)
            c.restoreState()

        draw_diamond(self.x, self.y, self.left_color)
        right_x = self.x + self.size + self.gap
        draw_diamond(right_x, self.y, self.right_color)

        if self.number:
            c.setFont(FONT, self.size//2)
            c.setFillColor(colors.white)
            num_y = self.y - (self.size // 4)
            c.drawCentredString(right_x, num_y, str(self.number))

        c.setFont(FONT, 30)
        c.setFillColor(self.right_color)
        text_x = right_x + self.size/2 + 4
        text_y = self.y - 2
        c.drawString(text_x, text_y, self.title)

# =========================
# 页码带箭头
# =========================
def draw_page_number(canvas, doc):
    page_num = str(doc.page).zfill(2)
    arrow_size = 6
    margin = 20
    w, h = A4

    canvas.setFont(FONT, 10)
    text_width = canvas.stringWidth(page_num, FONT, 10)
    x_text = w - margin - arrow_size - 4 - text_width
    y = margin
    x_arrow = w - margin - arrow_size

    # 页码文字
    canvas.setFillColor(colors.Color(0.45, 0.2, 0.6, 1))
    canvas.drawString(x_text, y, page_num)

    # 小箭头
    path = canvas.beginPath()
    path.moveTo(x_arrow + arrow_size, y)
    path.lineTo(x_arrow, y + arrow_size/2)
    path.lineTo(x_arrow + arrow_size, y + arrow_size)
    path.close()
    canvas.drawPath(path, stroke=0, fill=1)

# =========================
# 生成PDF
# =========================
def create_customer_info():
    doc = SimpleDocTemplate("Oncoseeing_基本信息页_最终版.pdf", pagesize=A4,
                            leftMargin=35, rightMargin=35, topMargin=30, bottomMargin=30)

    table_style = get_table_style()
    text_style = get_table_text_style()

    # 表格数据
    data = [
        ["项目", "详情"],  # 表头
        ["姓名", ""],
        ["年龄", ""],
        ["样本类型", "外周血"],
        ["样本编号", "2306052271"],
        ["样本ID", "211369970"],
        ["送样日期", ""],
        ["报告日期", ""],
        ["检测项目", "循环肿瘤DNA(ctDNA)中1893个癌症特异性突变检测"],
    ]

    data = [[Paragraph(cell, text_style) for cell in row] for row in data]
    table = Table(data, colWidths=[120, 340], rowHeights=36)
    table.setStyle(table_style)

    # 页面元素
    elements = []
    elements.append(Spacer(1, 80))

    # 菱形标题
    elements.append(SectionTitle("基本信息", number=1, x=0, y=0))
    elements.append(Spacer(1, 50))

    # 表格
    elements.append(table)

    # 构建PDF，绘制页码
    doc.build(elements, onFirstPage=draw_page_number, onLaterPages=draw_page_number)

    print("✅ 已生成：Oncoseeing_基本信息页_最终版.pdf")

# =========================
# 运行
# =========================
if __name__ == "__main__":
    create_customer_info()