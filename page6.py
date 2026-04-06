from reportlab.platypus import Table

from reportlab.platypus import Image
from PIL import Image as PILImage  # 用于获取原始尺寸
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
# 指向微软雅黑字体文件 (注意是 .ttc 格式)
FONT_PATH = FONT_DIR / "msyh.ttc"

if os.path.exists(FONT_PATH):
    # 注意：注册 TTC 字体需要指定索引，通常是 0
    pdfmetrics.registerFont(TTFont("MicrosoftYaHei", FONT_PATH, subfontIndex=0))
    FONT = "MicrosoftYaHei" # 设定字体名称
else:
    print("⚠️ 未找到微软雅黑字体，使用默认字体")
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
# =========================
# 菱形标题（优化间距版）
# =========================
class SectionTitle(Flowable):
    def __init__(self, title, number=None, x=0, y=0,
                 left_color=colors.Color(0.7, 0.6, 0.9),
                 right_color=colors.Color(0.34, 0.28, 0.63),
                 size=30, gap=-10):  # 👈 两个菱形更近：gap=-10
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
            c.rect(-self.size / 2, -self.size / 2, self.size, self.size, fill=1, stroke=0)
            c.restoreState()

        draw_diamond(self.x, self.y, self.left_color)
        right_x = self.x + self.size + self.gap
        draw_diamond(right_x, self.y, self.right_color)

        if self.number:
            c.setFont(FONT, self.size // 2)
            c.setFillColor(colors.white)
            num_y = self.y - (self.size // 4)
            c.drawCentredString(right_x, num_y, str(self.number))

        c.setFont(FONT, 30)
        c.setFillColor(self.right_color)

        # 👈 菱形 ↔ 文字拉开距离：+4 → +10
        text_x = right_x + self.size / 2 + 10

        text_y = self.y - 10  # 垂直居中（保持不变）
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

    # 左侧文字
    left_text = Paragraph(
        "Oncoseeing™检测采用国际前沿的液体活检技术，通过采集5mL外周血，"
        "提取循环肿瘤DNA(ctDNA)，利用高通量测序平台靶向检测1893个与癌症发生发展相关的特异性突变位点，"
        "并结合AI大数据模型进行风险评估与组织溯源。",
        subtitle_text_style
    )

    # 右侧图片（保持比例）
    img_path = BASE_DIR / "microscope2.png"
    if os.path.exists(img_path):
        pil_img = PILImage.open(img_path)
        orig_w, orig_h = pil_img.size

        max_width = 220  # 控制图片宽度（右侧区域）
        scale = max_width / orig_w
        img = Image(str(img_path), width=max_width, height=orig_h * scale)
    else:
        img = Paragraph("图片缺失", subtitle_text_style)

    # 两列布局
    content_table = Table(
        [[left_text, img]],
        colWidths=[(PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN) * 0.5,
                   (PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN) * 0.5]
    )

    content_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),  # 顶部对齐
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
    ]))

    info_data = [
        ["项目", "详情", "项目", "详情"],
        ["姓名", "测试患者", "性别", "男"],
        ["年龄", "45", "样本编号", "TEST001"],
        ["报告日期", "2025-01-05", "检测类型", "泛癌早筛"]
    ]

    # 转 Paragraph（关键！否则不会白字）
    info_data = [
        [Paragraph(cell, table_style_text) for cell in row]
        for row in info_data
    ]

    table = Table(info_data, colWidths=[70, 130, 70, 130], rowHeights=35)
    table.setStyle(get_table_style())

    elements = []
    elements.append(Spacer(1, 50))
    elements.append(SectionTitle("附录：检测技术与重要声明", number=6, x=0, y=0))
    elements.append(Spacer(1, 50))
    elements.append(SectionSubtitle("6.1 检测技术说明", x=0, y=0))

    # 5.1 核心建议下方插入文字
    page6_text = (
        "鉴于Oncoseeing™检测发现阳性突变信号，建议您尽早前往医院相关科室"
        "（如神经外科、妇科、消化内科/胃肠外科）进行专项检查。这有助于："
        "及时发现或排除早期病变。若检查无异常，可基本排除近期患相应癌症的风险。"
        "若检查发现异常，可尽早进行临床干预。"
    )

    elements.append(Spacer(1, 10))
    elements.append(content_table)
    elements.append(Spacer(1, 50))  # 可选：段落下方空白
    elements.append(SectionSubtitle("6.2 个体化健康管理建议", x=0, y=0))

    # 插入图片
    img_advice_path = BASE_DIR / "page6.png"
    if os.path.exists(img_advice_path):
        # 用 PIL 获取原始宽高
        pil_img = PILImage.open(img_advice_path)
        orig_width, orig_height = pil_img.size

        # 可用宽度
        max_width = PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN

        # 按比例计算高度
        scale = max_width / orig_width
        new_width = orig_width * scale
        new_height = orig_height * scale

        img = Image(str(img_advice_path), width=new_width, height=new_height)
        elements.append(img)
        elements.append(Spacer(1, 20))
    else:
        print(f"⚠️ 图片不存在: {img_advice_path}")

    elements.append(Spacer(1, 20))






    doc.build(elements)
    print("✅ 生成完成：final_pdf_example.pdf")