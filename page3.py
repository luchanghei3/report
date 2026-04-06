from reportlab.platypus import Image
from reportlab.platypus import SimpleDocTemplate, Spacer, Flowable
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib import colors
from reportlab.platypus import Paragraph
import os
from pathlib import Path
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
# =========================
# 全局配置
# =========================
BASE_DIR = Path(__file__).resolve().parent
FONT_DIR = Path("C:/Windows/Fonts")
FONT_PATH = FONT_DIR / "simhei.ttf"

if os.path.exists(FONT_PATH):
    pdfmetrics.registerFont(TTFont("simhei", FONT_PATH))
    FONT = "simhei"
else:
    FONT = "Helvetica"

# =========================
# 自定义文字样式
# =========================
subtitle_text_style = ParagraphStyle(
    name="SubtitleText",
    fontName=FONT,
    fontSize=12,
    leading=18,  # 行高
    textColor=colors.black,
    leftIndent=0,
    rightIndent=0,
    spaceAfter=10,  # 段后间距
)

from reportlab.platypus import Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet

def get_table_style():
    return TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.white),
        ("PADDING", (0, 0), (-1, -1), 6),
        ("BACKGROUND", (0, 1), (-1, -1), colors.Color(0.72, 0.58, 0.85, 1)),
        ("BACKGROUND", (0, 0), (-1, 0), colors.Color(0.45, 0.2, 0.6, 1)),
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
# 文字内容
subtitle_text_content = (
    "本次Oncoseeing™检测在您的外周血中发现了若干个癌症特异性突变，分别位于TP53和 PTEN基因上。"
    "基于AI大数据模型分析，这些突变提示您的体内可能存在具有癌化趋势的细胞，"
    "相关风险主要与脑胶质瘤、子宫内膜癌、结肠癌的成因相关。"
)

PAGE_WIDTH, PAGE_HEIGHT = A4
LEFT_MARGIN = 35
RIGHT_MARGIN = 35
TOP_MARGIN = 30
BOTTOM_MARGIN = 30

# =========================
# 菱形标题
# =========================
class SectionTitle(Flowable):
    def __init__(self, title, number=None, x=0, y=0,
                 left_color=colors.Color(0.7, 0.6, 0.9),
                 right_color=colors.Color(0.34, 0.28, 0.63),
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

        c.setFont(FONT, 30)  # 修改标题大小
        c.setFillColor(self.right_color)
        text_x = right_x + self.size/2 + 4
        text_y = self.y - 2
        c.drawString(text_x, text_y, self.title)

# =========================
# 二级标题（带箭头）
# =========================
class SectionSubtitle(Flowable):
    def __init__(self, title, x=0, y=0, arrow_color=colors.Color(0.45,0.2,0.6,1), font_size=13):
        super().__init__()
        self.title = title
        self.x = x
        self.y = y
        self.arrow_color = arrow_color
        self.font_size = font_size

    def draw(self):
        c = self.canv
        arrow_size = 10
        c.setFillColor(self.arrow_color)
        path = c.beginPath()
        path.moveTo(self.x, self.y)
        path.lineTo(self.x + arrow_size, self.y + arrow_size/2)
        path.lineTo(self.x, self.y + arrow_size)
        path.close()
        c.drawPath(path, stroke=0, fill=1)

        c.setFont(FONT, self.font_size)
        c.setFillColor(self.arrow_color)
        c.drawString(self.x + arrow_size + 4, self.y + 1, self.title)

# =========================
# 页码绘制
# =========================
def draw_page_number(canvas, doc):
    page_num = str(doc.page).zfill(2)  # 01, 02
    arrow_size = 6
    margin = 20
    w, h = A4

    canvas.setFont(FONT, 10)
    text_width = canvas.stringWidth(page_num, FONT, 10)
    x_text = w - margin - arrow_size - 4 - text_width
    y = margin
    x_arrow = w - margin - arrow_size

    # 绘制页码文字
    canvas.setFillColor(colors.Color(0.45,0.2,0.6,1))
    canvas.drawString(x_text, y, page_num)

    # 绘制向左箭头
    path = canvas.beginPath()
    path.moveTo(x_arrow + arrow_size, y)       # 右下
    path.lineTo(x_arrow, y + arrow_size/2)    # 左中
    path.lineTo(x_arrow + arrow_size, y + arrow_size)  # 右上
    path.close()
    canvas.drawPath(path, stroke=0, fill=1)

# =========================
# 底部图片
# =========================
def draw_bottom_full_image(canvas, doc):
    img_path = BASE_DIR / "dna.png"
    if not os.path.exists(img_path):
        print(f"⚠️ 图片不存在: {img_path}")
        return

    img_width = PAGE_WIDTH
    img_height = PAGE_HEIGHT * 0.35
    canvas.drawImage(
        str(img_path),
        0, 0,
        width=img_width,
        height=img_height,
        preserveAspectRatio=False,
        mask='auto'
    )

# =========================
# 生成 PDF
# =========================
if __name__ == "__main__":

    doc = SimpleDocTemplate(
        "final_pdf_example.pdf",
        pagesize=A4,
        leftMargin=LEFT_MARGIN,
        rightMargin=RIGHT_MARGIN,
        topMargin=TOP_MARGIN,
        bottomMargin=BOTTOM_MARGIN
    )

    table_style_text = get_table_text_style()

    mutation_data = [
        ["相关癌种", "基因", "碱基突变", "氨基酸突变", "突变比例"], # 表头
        ["脑胶质瘤、子宫内膜癌、结肠癌", "Tp53", "c.817C>T", "p.R273C", "0.084%"],
        ["脑胶质瘤、子宫内膜癌、结肠癌", "Tp53", "c.818G>A", "p.R273H", "0.056%"],
        ["子宫内膜癌", "PTEN", "c.389G>A", "p.R130Q", "0.054%"],
    ]

    # 转 Paragraph（关键！否则不会白字）
    mutation_data = [
        [Paragraph(cell, table_style_text) for cell in row]
        for row in mutation_data
    ]

    table = Table(mutation_data, colWidths=[70, 130, 70, 130], rowHeights=40)
    table.setStyle(get_table_style())

    elements = []
    elements.append(Spacer(1, 50))
    elements.append(SectionTitle("检测结果详情", number=3, x=0, y=0))
    elements.append(Spacer(1, 50))
    elements.append(SectionSubtitle("3.1 阳性发现特异性突变点详情", x=0, y=0))

    elements.append(Spacer(1, 10))
    elements.append(table)  # ⭐插入表格

    elements.append(Spacer(1, 50))  # 可选：段落下方空白
    elements.append(SectionSubtitle("3.2 基因功能简介", x=0, y=0))

    elements.append(Spacer(1, 20))

    # 插入图片
    img_path = BASE_DIR / "genedescription.png"
    if os.path.exists(img_path):
        img = Image(str(img_path), width=PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN, height=200)
        elements.append(img)
    else:
        print(f"⚠️ 图片不存在: {img_path}")

    elements.append(Spacer(1, 30))


    # build 时同时绘制底部图片和页码
    def on_page(canvas, doc):
        draw_bottom_full_image(canvas, doc)
        draw_page_number(canvas, doc)

    doc.build(elements, onFirstPage=on_page, onLaterPages=on_page)
    print("✅ 生成完成：final_pdf_example.pdf")