from reportlab.platypus import Table, TableStyle, Paragraph, Spacer, Image, SimpleDocTemplate, Flowable
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
import os
from pathlib import Path

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

PAGE_WIDTH, PAGE_HEIGHT = A4
LEFT_MARGIN = 35
RIGHT_MARGIN = 35
TOP_MARGIN = 30
BOTTOM_MARGIN = 30

# =========================
# 文字样式
# =========================
subtitle_text_style_4 = ParagraphStyle(
    name="SubtitleText",
    fontName=FONT,
    fontSize=12,
    leading=20,
    firstLineIndent=24,
    alignment=4,  # 两端对齐
    spaceAfter=10,
)

def get_table_text_style():
    return ParagraphStyle(
        name="TableCell",
        fontName=FONT,
        fontSize=11,
        leading=14,
        alignment=1,  # 居中
        textColor=colors.white
    )

table_style_text = get_table_text_style()

# =========================
# 小表格（癌症表格）
# =========================
def create_cancer_table(title, desc):
    table_style_text = get_table_text_style()
    data = [
        [Paragraph(title, table_style_text)],
        [Paragraph(desc, table_style_text)]
    ]
    t = Table(data, colWidths=[(PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN) / 3],
              rowHeights=[30, 50])  # 固定行高
    t.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.white),
        ("BACKGROUND", (0,0), (-1,0), colors.Color(0.45,0.2,0.6,1)),  # 表头深紫色
        ("BACKGROUND", (0,1), (-1,1), colors.Color(0.72,0.58,0.85,1)), # 内容浅紫色
        ("TEXTCOLOR", (0,0), (-1,-1), colors.white),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("PADDING", (0,0), (-1,-1), 6),
    ]))
    return t

# 三个癌症表
table1 = create_cancer_table("脑胶质瘤相关位点2个",
                             "对比该癌种患者突变中位数（9个），您的风险信号处于早期阶段。")
table2 = create_cancer_table("子宫内膜癌相关位点3个",
                             "对比该癌种患者突变中位数（4个），您的风险信号需引起关注。")
table3 = create_cancer_table("结肠癌相关位点2个",
                             "对比该癌种患者突变中位数（4个），您的风险信号处于早期阶段。")

# 并排三列
row_table = Table([[table1, table2, table3]],
                  colWidths=[(PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN)/3]*3)
row_table.setStyle(TableStyle([
    ("VALIGN", (0,0), (-1,-1), "TOP"),
    ("ALIGN", (0,0), (-1,-1), "CENTER"),
    ("LEFTPADDING", (0,0), (-1,-1), 0),
    ("RIGHTPADDING", (0,0), (-1,-1), 0),
]))

# =========================
# info 表格样式
# =========================
def get_table_style():
    return TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.white),
        ("BACKGROUND", (0,0), (-1,0), colors.Color(0.45,0.2,0.6,1)),  # 表头深紫
        ("BACKGROUND", (0,1), (-1,-1), colors.Color(0.72,0.58,0.85,1)), # 内容浅紫
        ("TEXTCOLOR", (0,0), (-1,-1), colors.white),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("PADDING", (0,0), (-1,-1), 6),
    ])

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

        c.setFont(FONT, 30)
        c.setFillColor(self.right_color)
        text_x = right_x + self.size/2 + 4
        text_y = self.y - 2
        c.drawString(text_x, text_y, self.title)

# =========================
# 二级标题
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
# 底部图片
# =========================
def draw_bottom_full_image(canvas, doc):
    img_path = BASE_DIR / "doctor.png"
    if not os.path.exists(img_path):
        print(f"⚠️ 图片不存在: {img_path}")
        return
    img_width = PAGE_WIDTH
    img_height = PAGE_HEIGHT * 0.35
    canvas.drawImage(str(img_path), 0, 0, width=img_width, height=img_height, preserveAspectRatio=False, mask='auto')

# =========================
# PDF生成
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

    # info表格内容
    info_data = [
        ["项目", "详情", "项目", "详情"],
        ["姓名", "测试患者", "性别", "男"],
        ["年龄", "45", "样本编号", "TEST001"],
        ["报告日期", "2025-01-05", "检测类型", "泛癌早筛"]
    ]
    info_data = [[Paragraph(cell, table_style_text) for cell in row] for row in info_data]
    table = Table(info_data, colWidths=[70, 130, 70, 130], rowHeights=[35]*len(info_data))
    table.setStyle(get_table_style())

    # PDF元素
    elements = []
    elements.append(Spacer(1, 50))
    elements.append(SectionTitle("阳性发现深度解读", number=4, x=0, y=0))
    elements.append(Spacer(1, 50))
    elements.append(SectionSubtitle("4.1 风险评估解读", x=0, y=0))
    elements.append(Spacer(1, 10))
    elements.append(Paragraph("您的Oncoseeing™检测结果中，与三种癌症相关的突变位点数量均低于该癌种患者群体的中位数水平，但已明确提示存在癌化风险信号。", subtitle_text_style_4))
    elements.append(Spacer(1, 20))
    elements.append(row_table)
    elements.append(Spacer(1, 30))
    elements.append(SectionSubtitle("4.2 变异细胞数量估算", x=0, y=0))
    elements.append(Spacer(1, 10))
    elements.append(Paragraph("根据突变比例与细胞数量的研究模型估算，您体内携带上述特异性突变的变异细胞数量约<br/>在 10^7 ~ 10^8 个。此阶段通常对应癌前病变或极早期阶段，影像学检查可能难以发现，正是进行主动干预的关键窗口期。", subtitle_text_style_4))
    elements.append(Spacer(1, 20))
    elements.append(SectionSubtitle("4.3 临床意义与数据库比对", x=0, y=0))
    elements.append(Spacer(1, 10))
    elements.append(Paragraph("上述突变在OBCD等大型肿瘤突变数据库的相应癌症患者中均有记录，但在大规模正常人群样本库中不存在，属于“癌症特异性突变”。Oncoseeing™模型溯源算法提示，突变来源的可能性为子宫内膜或结肠。", subtitle_text_style_4))
    elements.append(Spacer(1, 10))

    # 插入图片
    img_path = BASE_DIR / "genedescription.png"
    if os.path.exists(img_path):
        img = Image(str(img_path), width=PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN, height=200)
        elements.append(img)
    else:
        print(f"⚠️ 图片不存在: {img_path}")
    elements.append(Spacer(1, 30))

    # build 时绘制底部图片和页码
    def on_page(canvas, doc):
        draw_bottom_full_image(canvas, doc)
        draw_page_number(canvas, doc)

    doc.build(elements, onFirstPage=on_page, onLaterPages=on_page)
    print("✅ 生成完成：final_pdf_example.pdf")