from reportlab.platypus import SimpleDocTemplate, Spacer, Flowable, Table, TableStyle, PageBreak
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib import colors
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
import os
from pathlib import Path

# =========================
# 全局配置
# =========================
# BASE_DIR = Path(__file__).resolve().parent
# FONT_DIR = Path("C:/Windows/Fonts")
# FONT_PATH = FONT_DIR / "simhei.ttf"
#
# if os.path.exists(FONT_PATH):
#     pdfmetrics.registerFont(TTFont("simhei", FONT_PATH))
#     FONT = "simhei"
# else:
#     FONT = "Helvetica"

# =========================
# 全局配置 - 修复版
# =========================
BASE_DIR = Path(__file__).resolve().parent
FONT_DIR = Path("C:/Windows/Fonts")
# 指向微软雅黑字体文件 (注意是 .ttc 格式)
FONT_PATH = FONT_DIR / "msyh.ttc"

if os.path.exists(FONT_PATH):
    # 注意：注册 TTC 字体需要指定索引，通常是 0
    pdfmetrics.registerFont(TTFont("MicrosoftYaHei", FONT_PATH, subfontIndex=0))
    FONT = "MicrosoftYaHei" # 设定字体名称
else:
    print("⚠️ 未找到微软雅黑字体，使用默认字体")
    FONT = "Helvetica"

PAGE_WIDTH, PAGE_HEIGHT = A4
LEFT_MARGIN = 35
RIGHT_MARGIN = 35
TOP_MARGIN = 30
BOTTOM_MARGIN = 30

# ======================
# 底部图片高度 + 内容安全区（真正生效的分页线）
# ======================
BOTTOM_IMG_HEIGHT = PAGE_HEIGHT * 0.35
SAFE_BOTTOM = BOTTOM_MARGIN + BOTTOM_IMG_HEIGHT

# =========================
# 文字样式
# =========================
subtitle_text_style = ParagraphStyle(
    name="SubtitleText",
    fontName=FONT,
    fontSize=12,
    leading=18,
    textColor=colors.black,
    leftIndent=0,
    rightIndent=0,
    spaceAfter=10,
)

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


# =========================
# 自定义表格：分页后自动加顶部间距
# =========================
# =========================
# 标题组件（菱形）
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
        self.width = 300
        self.height = 40

    def wrap(self, aw, ah):
        return self.width, self.height

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
        self.width = 300
        self.height = 20

    def wrap(self, aw, ah):
        return self.width, self.height

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
# 页码
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

    canvas.setFillColor(colors.Color(0.45,0.2,0.6,1))
    canvas.drawString(x_text, y, page_num)

    path = canvas.beginPath()
    path.moveTo(x_arrow + arrow_size, y)
    path.lineTo(x_arrow, y + arrow_size/2)
    path.lineTo(x_arrow + arrow_size, y + arrow_size)
    path.close()
    canvas.drawPath(path, stroke=0, fill=1)

# =========================
# 底部图片（每页都显示）
# =========================
def draw_bottom_full_image(canvas, doc):
    img_path = BASE_DIR / "microscope.png"
    if not os.path.exists(img_path):
        print(f"⚠️ 图片不存在: {img_path}")
        return

    canvas.drawImage(
        str(img_path),
        0, 0,
        width=PAGE_WIDTH,
        height=BOTTOM_IMG_HEIGHT,
        preserveAspectRatio=False,
        mask='auto'
    )

# =========================
# 每页统一样式
# =========================
def on_page(canvas, doc):
    draw_bottom_full_image(canvas, doc)
    draw_page_number(canvas, doc)



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
        bottomMargin=SAFE_BOTTOM,  # 👈 核心修复：把底部边距直接设为图片高度
    )

    # 内容
    subtitle_text_content = (
        "本次Oncoseeing™检测在您的外周血中发现了若干个癌症特异性突变，分别位于TP53和 PTEN基因上。"
        "基于AI大数据模型分析，这些突变提示您的体内可能存在具有癌化趋势的细胞，"
        "相关风险主要与脑胶质瘤、子宫内膜癌、结肠癌的成因相关。"
    )

    # 表格样式
    table_style_text = get_table_text_style()

    data = [
        ["编号", "患癌风险组织", "检测评估结果", "癌化风险"],
        ["1", "脑胶质细胞", "检出特异性突变点2个", "低"],
        ["2", "子宫内膜", "检出特异性突变点3个", "低"],
        ["3", "结肠", "检出特异性突变点2个", "低"],
        ["4", "乳腺", "未检出特异性突变点", "阴"],
        ["5", "肺", "未检出特异性突变点", "阴"],
        ["6", "胆道", "未检出特异性突变点", "阴"],
        ["7", "骨", "未检出特异性突变点", "阴"],
        ["8", "宫颈", "未检出特异性突变点", "阴"],
        ["9", "造血系统", "未检出特异性突变点", "阴"],
        ["10", "淋巴组织", "未检出特异性突变点", "阴"],
        ["11", "肾", "未检出特异性突变点", "阴"],
        ["12", "肝", "未检出特异性突变点", "阴"],
        ["13", "尿路", "未检出特异性突变点", "阴"],
        ["14", "食管", "未检出特异性突变点", "阴"],
        ["15", "卵巢", "未检出特异性突变点", "阴"],
        ["16", "胰腺", "未检出特异性突变点", "阴"],
        ["17", "皮肤", "未检出特异性突变点", "阴"],
        ["18", "胃", "未检出特异性突变点", "阴"],
        ["19", "甲状腺", "未检出特异性突变点", "阴"],
        ["20", "上呼吸道", "未检出特异性突变点", "阴"],

    ]

    data = [[Paragraph(cell, table_style_text) for cell in row] for row in data]

    # 表格自动分页 + 每页重复表头
    table = Table(
        data,
        colWidths=[70, 160, 200, 70],
        rowHeights=28,
        repeatRows=1,
    )

    # ==============================
    table.setStyle(get_table_style())

    # 页面元素
    elements = []
    elements.append(Spacer(1, 50))
    elements.append(SectionTitle("检测结果综合评估", number=2, x=0, y=0))
    elements.append(Spacer(1, 30))

    elements.append(SectionSubtitle("2.1 综合结论摘要", x=0, y=0))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph(subtitle_text_content, subtitle_text_style))
    elements.append(Spacer(1, 30))

    elements.append(SectionSubtitle("2.2 风险评估总览", x=0, y=0))
    elements.append(Spacer(1, 20))
    elements.append(table)

    # 构建
    doc.build(elements, onFirstPage=on_page, onLaterPages=on_page)
    print("✅ PDF 生成完成：final_pdf_example.pdf")