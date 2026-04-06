"""统一生成 Oncoseeing page1-6，并拼接 coverpage/lastpage。

本脚本把原先分散在 customerinfo.py、page1_new.py、page3.py、page4.py、page5.py、page6.py
中的核心页面内容集中到一个入口，统一样式、页码和构建流程。

用法:
    python merge_report.py

依赖:
    pip install reportlab pypdf pillow
"""

from __future__ import annotations

from pathlib import Path
import os

try:
    from reportlab.platypus import (
        SimpleDocTemplate,
        Table,
        TableStyle,
        Paragraph,
        Spacer,
        Flowable,
        Image,
        PageBreak,
    )
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase import pdfmetrics
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
except Exception as exc:  # noqa: BLE001
    raise RuntimeError("缺少 reportlab，请先执行: pip install reportlab") from exc

try:
    from PIL import Image as PILImage
except Exception as exc:  # noqa: BLE001
    raise RuntimeError("缺少 pillow，请先执行: pip install pillow") from exc

try:
    from pypdf import PdfWriter
except Exception as exc:  # noqa: BLE001
    raise RuntimeError("缺少 pypdf，请先执行: pip install pypdf") from exc


BASE_DIR = Path(__file__).resolve().parent
PAGE_WIDTH, PAGE_HEIGHT = A4
LEFT_MARGIN = 35
RIGHT_MARGIN = 35
TOP_MARGIN = 30
BOTTOM_MARGIN = 30


def setup_font() -> str:
    yahei = Path("C:/Windows/Fonts/msyh.ttc")
    simhei = Path("C:/Windows/Fonts/simhei.ttf")
    if os.path.exists(yahei):
        pdfmetrics.registerFont(TTFont("ReportFont", str(yahei), subfontIndex=0))
        return "ReportFont"
    if os.path.exists(simhei):
        pdfmetrics.registerFont(TTFont("ReportFont", str(simhei)))
        return "ReportFont"
    return "Helvetica"


FONT = setup_font()


def purple() -> colors.Color:
    return colors.Color(0.45, 0.2, 0.6, 1)


def light_purple() -> colors.Color:
    return colors.Color(0.72, 0.58, 0.85, 1)


class SectionTitle(Flowable):
    def __init__(self, title: str, number: int, size: int = 28, gap: int = -8):
        super().__init__()
        self.title = title
        self.number = number
        self.size = size
        self.gap = gap
        self.width = PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN
        self.height = 42

    def wrap(self, aw, ah):
        return self.width, self.height

    def draw(self):
        c = self.canv
        x = 0
        y = 20

        def draw_diamond(px, py, color):
            c.setFillColor(color)
            c.saveState()
            c.translate(px, py)
            c.rotate(45)
            c.rect(-self.size / 2, -self.size / 2, self.size, self.size, fill=1, stroke=0)
            c.restoreState()

        draw_diamond(x + self.size / 2, y, light_purple())
        right_x = x + self.size + self.gap
        draw_diamond(right_x, y, colors.Color(0.34, 0.28, 0.63, 1))

        c.setFont(FONT, self.size // 2)
        c.setFillColor(colors.white)
        c.drawCentredString(right_x, y - (self.size // 4), str(self.number))

        c.setFont(FONT, 26)
        c.setFillColor(colors.Color(0.34, 0.28, 0.63, 1))
        c.drawString(right_x + self.size / 2 + 10, y - 8, self.title)


class SectionSubtitle(Flowable):
    def __init__(self, title: str, font_size: int = 13):
        super().__init__()
        self.title = title
        self.font_size = font_size
        self.width = PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN
        self.height = 20

    def wrap(self, aw, ah):
        return self.width, self.height

    def draw(self):
        c = self.canv
        x = 0
        y = 6
        arrow_size = 10

        c.setFillColor(purple())
        path = c.beginPath()
        path.moveTo(x, y)
        path.lineTo(x + arrow_size, y + arrow_size / 2)
        path.lineTo(x, y + arrow_size)
        path.close()
        c.drawPath(path, stroke=0, fill=1)

        c.setFont(FONT, self.font_size)
        c.setFillColor(purple())
        c.drawString(x + arrow_size + 4, y + 1, self.title)


def draw_page_number(canvas, doc):
    page_num = str(doc.page).zfill(2)
    arrow_size = 6
    margin = 20

    canvas.setFont(FONT, 10)
    text_width = canvas.stringWidth(page_num, FONT, 10)
    x_text = PAGE_WIDTH - margin - arrow_size - 4 - text_width
    y = margin
    x_arrow = PAGE_WIDTH - margin - arrow_size

    canvas.setFillColor(purple())
    canvas.drawString(x_text, y, page_num)

    path = canvas.beginPath()
    path.moveTo(x_arrow + arrow_size, y)
    path.lineTo(x_arrow, y + arrow_size / 2)
    path.lineTo(x_arrow + arrow_size, y + arrow_size)
    path.close()
    canvas.drawPath(path, stroke=0, fill=1)


def draw_bottom_image(canvas, image_name: str, height_ratio: float = 0.35):
    img_path = BASE_DIR / image_name
    if not img_path.exists():
        return
    canvas.drawImage(
        str(img_path),
        0,
        0,
        width=PAGE_WIDTH,
        height=PAGE_HEIGHT * height_ratio,
        preserveAspectRatio=False,
        mask="auto",
    )


def on_page(canvas, doc):
    # 保留原 page3/4/5 的底部 PNG 逻辑
    bottom_map = {
        3: "dna.png",     # page3
        4: "doctor.png",  # page4
        5: "body.png",    # page5
    }
    image_name = bottom_map.get(doc.page)
    if image_name:
        draw_bottom_image(canvas, image_name)
    draw_page_number(canvas, doc)


def table_text_style() -> ParagraphStyle:
    return ParagraphStyle(
        name="TableCell",
        fontName=FONT,
        fontSize=11,
        leading=14,
        alignment=1,
        textColor=colors.white,
    )


def body_style() -> ParagraphStyle:
    return ParagraphStyle(
        name="BodyText",
        fontName=FONT,
        fontSize=12,
        leading=20,
        textColor=colors.black,
    )


def body_justify_style() -> ParagraphStyle:
    return ParagraphStyle(
        name="BodyJustify",
        fontName=FONT,
        fontSize=12,
        leading=20,
        firstLineIndent=24,
        alignment=4,
        textColor=colors.black,
    )


def common_table_style() -> TableStyle:
    return TableStyle(
        [
            ("GRID", (0, 0), (-1, -1), 0.5, colors.white),
            ("PADDING", (0, 0), (-1, -1), 6),
            ("BACKGROUND", (0, 0), (-1, 0), purple()),
            ("BACKGROUND", (0, 1), (-1, -1), light_purple()),
            ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]
    )


def add_scaled_image(elements: list, image_name: str, max_width: float, max_height: float | None = None):
    img_path = BASE_DIR / image_name
    if not img_path.exists():
        elements.append(Paragraph(f"图片缺失：{image_name}", body_style()))
        return

    pil_img = PILImage.open(img_path)
    orig_w, orig_h = pil_img.size
    scale = max_width / orig_w
    if max_height is not None and orig_h * scale > max_height:
        scale = min(scale, max_height / orig_h)

    elements.append(Image(str(img_path), width=orig_w * scale, height=orig_h * scale))


def build_unified_pages(output_pdf: Path):
    doc = SimpleDocTemplate(
        str(output_pdf),
        pagesize=A4,
        leftMargin=LEFT_MARGIN,
        rightMargin=RIGHT_MARGIN,
        topMargin=TOP_MARGIN,
        bottomMargin=BOTTOM_MARGIN,
    )

    tstyle = table_text_style()
    bst = body_style()
    jst = body_justify_style()
    elements = []

    # -------- Page1 (from customerinfo.py) --------
    info1 = [
        ["项目", "详情"],
        ["姓名", ""],
        ["年龄", ""],
        ["样本类型", "外周血"],
        ["样本编号", "2306052271"],
        ["样本ID", "211369970"],
        ["送样日期", ""],
        ["报告日期", ""],
        ["检测项目", "循环肿瘤DNA(ctDNA)中1893个癌症特异性突变检测"],
    ]
    info1 = [[Paragraph(cell, tstyle) for cell in row] for row in info1]

    elements += [
        Spacer(1, 50),
        SectionTitle("基本信息", number=1),
        Spacer(1, 35),
        Table(info1, colWidths=[120, 340], rowHeights=36, style=common_table_style()),
        PageBreak(),
    ]

    # -------- Page2 (from page1_new.py) --------
    summary_text = (
        "本次Oncoseeing™检测在您的外周血中发现了若干个癌症特异性突变，分别位于TP53和 PTEN基因上。"
        "基于AI大数据模型分析，这些突变提示您的体内可能存在具有癌化趋势的细胞，"
        "相关风险主要与脑胶质瘤、子宫内膜癌、结肠癌的成因相关。"
    )
    risk_data = [
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
    risk_data = [[Paragraph(cell, tstyle) for cell in row] for row in risk_data]
    risk_table = Table(
        risk_data,
        # 调宽最右列，避免“癌化风险”显示过窄
        colWidths=[60, 170, 220, 75],
        rowHeights=24,
        repeatRows=1,
        style=common_table_style(),
    )

    elements += [
        Spacer(1, 25),
        SectionTitle("检测结果综合评估", number=2),
        Spacer(1, 20),
        SectionSubtitle("2.1 综合结论摘要"),
        Spacer(1, 10),
        Paragraph(summary_text, bst),
        Spacer(1, 15),
        SectionSubtitle("2.2 风险评估总览"),
        Spacer(1, 10),
        risk_table,
        PageBreak(),
    ]

    # -------- Page3 (from page3.py) --------
    mutation_data = [
        ["相关癌种", "基因", "碱基突变", "氨基酸突变", "突变比例"],
        ["脑胶质瘤、子宫内膜癌、结肠癌", "Tp53", "c.817C>T", "p.R273C", "0.084%"],
        ["脑胶质瘤、子宫内膜癌、结肠癌", "Tp53", "c.818G>A", "p.R273H", "0.056%"],
        ["子宫内膜癌", "PTEN", "c.389G>A", "p.R130Q", "0.054%"],
    ]
    mutation_data = [[Paragraph(cell, tstyle) for cell in row] for row in mutation_data]

    elements += [
        Spacer(1, 40),
        SectionTitle("检测结果详情", number=3),
        Spacer(1, 30),
        SectionSubtitle("3.1 阳性发现特异性突变点详情"),
        Spacer(1, 10),
        Table(
            mutation_data,
            colWidths=[130, 70, 120, 90, 80],
            rowHeights=32,
            style=common_table_style(),
        ),
        Spacer(1, 25),
        SectionSubtitle("3.2 基因功能简介"),
        Spacer(1, 10),
    ]
    add_scaled_image(elements, "genedescription.png", PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN, max_height=260)
    elements += [Spacer(1, 10), PageBreak()]

    # -------- Page4 (from page4.py) --------
    elements += [
        Spacer(1, 35),
        SectionTitle("阳性发现深度解读", number=4),
        Spacer(1, 30),
        SectionSubtitle("4.1 风险评估解读"),
        Spacer(1, 10),
        Paragraph(
            "您的Oncoseeing™检测结果中，与三种癌症相关的突变位点数量均低于该癌种患者群体的中位数水平，但已明确提示存在癌化风险信号。",
            jst,
        ),
        Spacer(1, 14),
    ]

    def mini_table(title: str, desc: str) -> Table:
        data = [[Paragraph(title, tstyle)], [Paragraph(desc, tstyle)]]
        tb = Table(data, colWidths=[170], rowHeights=[30, 64], style=common_table_style())
        return tb

    row = Table(
        [[
            mini_table("脑胶质瘤相关位点2个", "对比该癌种患者突变中位数（9个），您的风险信号处于早期阶段。"),
            mini_table("子宫内膜癌相关位点3个", "对比该癌种患者突变中位数（4个），您的风险信号需引起关注。"),
            mini_table("结肠癌相关位点2个", "对比该癌种患者突变中位数（4个），您的风险信号处于早期阶段。"),
        ]],
        colWidths=[170, 170, 170],
    )
    row.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ])
    )

    elements += [
        row,
        Spacer(1, 18),
        SectionSubtitle("4.2 变异细胞数量估算"),
        Spacer(1, 8),
        Paragraph(
            "根据突变比例与细胞数量的研究模型估算，您体内携带上述特异性突变的变异细胞数量约在 10^7 ~ 10^8 个。"
            "此阶段通常对应癌前病变或极早期阶段，影像学检查可能难以发现，正是进行主动干预的关键窗口期。",
            jst,
        ),
        Spacer(1, 12),
        SectionSubtitle("4.3 临床意义与数据库比对"),
        Spacer(1, 8),
        Paragraph(
            "上述突变在OBCD等大型肿瘤突变数据库的相应癌症患者中均有记录，但在大规模正常人群样本库中不存在，属于“癌症特异性突变”。"
            "Oncoseeing™模型溯源算法提示，突变来源的可能性为子宫内膜或结肠。",
            jst,
        ),
        PageBreak(),
    ]

    # -------- Page5 (from page5.py) --------
    elements += [
        Spacer(1, 40),
        SectionTitle("检测结果详情", number=5),
        Spacer(1, 30),
        SectionSubtitle("5.1 核心建议"),
        Spacer(1, 10),
        Paragraph(
            "鉴于Oncoseeing™检测发现阳性突变信号，建议您尽早前往医院相关科室（如神经外科、妇科、消化内科/胃肠外科）进行专项检查。"
            "这有助于及时发现或排除早期病变。若检查无异常，可基本排除近期患相应癌症的风险；若检查发现异常，可尽早进行临床干预。",
            bst,
        ),
        Spacer(1, 25),
        SectionSubtitle("5.2 个体化健康管理建议"),
        Spacer(1, 10),
    ]
    add_scaled_image(elements, "advice.png", PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN, max_height=350)
    elements += [PageBreak()]

    # -------- Page6 (from page6.py) --------
    elements += [
        Spacer(1, 40),
        SectionTitle("附录：检测技术与重要声明", number=6),
        Spacer(1, 30),
        SectionSubtitle("6.1 检测技术说明"),
        Spacer(1, 10),
    ]

    left_text = Paragraph(
        "Oncoseeing™检测采用国际前沿的液体活检技术，通过采集5mL外周血，提取循环肿瘤DNA(ctDNA)，"
        "利用高通量测序平台靶向检测1893个与癌症发生发展相关的特异性突变位点，并结合AI大数据模型进行风险评估与组织溯源。",
        bst,
    )

    right_cell: list = []
    add_scaled_image(right_cell, "microscope2.png", max_width=220, max_height=180)
    right_obj = right_cell[0]

    two_col = Table(
        [[left_text, right_obj]],
        colWidths=[(PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN) * 0.55, (PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN) * 0.45],
    )
    two_col.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ])
    )

    elements += [
        two_col,
        Spacer(1, 20),
        SectionSubtitle("6.2 重要声明"),
        Spacer(1, 10),
    ]
    add_scaled_image(elements, "page6.png", PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN, max_height=260)

    doc.build(elements, onFirstPage=on_page, onLaterPages=on_page)


def merge_full_report(content_pdf: Path, output_pdf: Path):
    cover = BASE_DIR / "coverpage.pdf"
    last = BASE_DIR / "lastpage.pdf"

    missing = [str(p.name) for p in (cover, last) if not p.exists()]
    if missing:
        raise FileNotFoundError(f"缺少输入文件: {', '.join(missing)}")

    # 🔥 新版合并方式
    merger = PdfWriter()
    merger.append(str(cover))
    merger.append(str(content_pdf))
    merger.append(str(last))
    with open(output_pdf, "wb") as f:
        merger.write(f)
    merger.close()

def main():
    content_pdf = BASE_DIR / "report_page1_6_unified.pdf"
    final_pdf = BASE_DIR / "Oncoseeing_完整报告_合并版.pdf"

    build_unified_pages(content_pdf)
    merge_full_report(content_pdf, final_pdf)

    print(f"✅ 已生成内容页: {content_pdf}")
    print(f"✅ 已生成最终合并报告: {final_pdf}")


if __name__ == "__main__":
    main()
