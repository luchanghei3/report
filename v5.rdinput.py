import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path
import os
import argparse
import tempfile

from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image,
    Frame, PageTemplate, NextPageTemplate, BaseDocTemplate
)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import inch
from reportlab.lib.utils import simpleSplit


def purple():
    return colors.Color(0.45, 0.2, 0.6, 1)


def light_purple():
    return colors.Color(0.72, 0.58, 0.85, 1)



# =========================
# 全局配置：解决Matplotlib中文显示问题
# =========================
plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']  # 优先使用黑体，兼容中文
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题
plt.rcParams['figure.autolayout'] = False  # 禁用自动布局，彻底避免tight_layout警告



# ========== 新增：定义品牌化色彩体系（蓝白+浅灰+淡黄绿） ==========
THEME_COLORS = {
    "brand_blue": colors.Color(0.15, 0.35, 0.65),  # 品牌蓝（沉稳不鲜艳）
    "light_gray": colors.Color(0.92, 0.94, 0.96),   # 浅灰（背景/辅助）
    "pale_yellow_green": colors.Color(0.95, 0.98, 0.85),  # 淡黄绿（提示/强调）
    "white": colors.white,          # 基础白色
    "text_black": colors.Color(0.1, 0.1, 0.1),      # 深灰文字（替代纯黑更舒缓）
    "risk_low": colors.Color(0.2, 0.6, 0.3),        # 低风险（淡绿，不刺眼）
    "risk_medium": colors.Color(0.8, 0.7, 0.2),     # 中风险（淡黄）
    "risk_high": colors.Color(0.7, 0.2, 0.2),       # 高风险（暗红，降低饱和度）
    "border_gray": colors.Color(0.85, 0.87, 0.89)   # 边框浅灰
}
# ===============================================================
# =========================
# 风险计算与可视化函数（优化箭头清晰度+低比例显示）
# =========================
def calculate_cancer_risk(specific_sites, total_sites, alpha=80, beta=0.01):
    """计算调整型Logistic癌化风险"""
    if total_sites == 0:
        return 0.0
    p = specific_sites / total_sites
    # 原始Sigmoid
    sigmoid = 1 / (1 + np.exp(-alpha * (p - beta)))
    # Sigmoid在p=0的偏移
    sigmoid_min = 1 / (1 + np.exp(alpha * beta))
    # 归一化到[0,1]
    risk = (sigmoid - sigmoid_min) / (1 - sigmoid_min)
    return risk



def generate_risk_bar_image(risk_value, risk_text, temp_dir, table_col_width=110):
    """
    透明背景版：
    1. 红绿条背景完全透明
    2. 箭头顶点落红绿条上沿，上沿对齐原位置
    3. 适配表格，箭头不扁
    """
    temp_file = tempfile.NamedTemporaryFile(
        suffix='.png',  # 必须用PNG格式（支持透明）
        dir=temp_dir,
        delete=False
    )
    temp_path = temp_file.name
    temp_file.close()

    # 画布尺寸：保证高度足够，避免表格压缩变形
    fig_width = table_col_width / 72
    fig_height = 40 / 72
    # ========== 关键1：设置画布背景为透明 ==========
    fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=300, facecolor='none')  # 画布透明
    ax.patch.set_alpha(0.0)  # 坐标轴区域也透明

    # 红绿条位置（固定）
    bar_x_start = 0.1
    bar_x_end = 0.9
    bar_y_start = 0.3   # 红绿条下沿
    bar_y_end = 0.5     # 红绿条上沿（箭头顶点要落在这）
    gradient = np.linspace(0, 1, 2000).reshape(1, -1)
    
    # 绘制渐变条（保持原有颜色，背景透明）
    im = ax.imshow(
        gradient,
        aspect='auto',
        cmap='RdYlGn_r',
        extent=[bar_x_start, bar_x_end, bar_y_start, bar_y_end],
        interpolation='bicubic'
    )
    # 渐变条本身不透明，只让背景透明（关键）
    im.set_alpha(1.0)

    # 箭头核心坐标计算
    bar_length = bar_x_end - bar_x_start
    arrow_x = bar_x_start + risk_value * bar_length
    arrow_top_y = 0.7  # 箭头上沿位置
    arrow_tip_y = bar_y_end  # 箭头顶点落红绿条上沿
    triangle_w = 0.05    # 箭头宽度
    triangle_h = arrow_top_y - arrow_tip_y  # 箭头高度

    # 限制箭头不超出红绿条范围
    arrow_x = max(bar_x_start + triangle_w, min(bar_x_end - triangle_w, arrow_x))

    # 绘制箭头（黑色实心）
    triangle_x = [arrow_x, arrow_x - triangle_w, arrow_x + triangle_w]
    triangle_y = [arrow_tip_y, arrow_top_y, arrow_top_y]
    ax.fill(triangle_x, triangle_y, 'black', edgecolor='none')

    # 隐藏坐标轴，保持干净
    ax.set_xticks([])
    ax.set_yticks([])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    for spine in ax.spines.values():
        spine.set_visible(False)

    # ========== 关键2：保存为透明背景的PNG ==========
    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    plt.savefig(
        temp_path,
        format='png',
        bbox_inches='tight',
        pad_inches=0,
        facecolor='none',  # 画布背景透明
        edgecolor='none',
        dpi=300,
        pil_kwargs={'quality': 100, 'transparent': True}  # 强制透明
    )
    plt.close(fig)

    return temp_path
# =========================
# 性别对应癌种过滤函数（通用，适配二维列表和DataFrame）
# =========================
def filter_cancer_by_gender(cancer_data, gender):
    """
    根据性别过滤癌种数据（男性移除乳腺癌，女性移除前列腺癌）
    :param cancer_data: 原始癌种数据（二维列表 / pandas DataFrame）
    :param gender: 性别（男/女）
    :return: 过滤后的癌种数据（保持输入数据格式）
    """
    if gender not in ["男", "女"]:
        return cancer_data

    # 处理pandas DataFrame格式（Excel读取后的数据）
    if isinstance(cancer_data, pd.DataFrame):
        # 筛选所有包含癌种名称的列（模糊匹配，兼容不同Excel表头）
        cancer_columns = [col for col in cancer_data.columns if
                          any(keyword in str(col) for keyword in ["癌种", "癌症", "肿瘤", "组织"])]
        if not cancer_columns:
            # 若未找到明确癌种列，默认按第一列筛选
            cancer_columns = [cancer_data.columns[0]]

        # 构建过滤条件
        filter_condition = pd.Series([True] * len(cancer_data))
        if gender == "男":
            # 男性：移除包含乳腺癌相关的行
            for col in cancer_columns:
                filter_condition = filter_condition & ~cancer_data[col].astype(str).str.contains("乳腺癌|乳腺导管腺癌",
                                                                                                 na=False)
        else:
            # 女性：移除包含前列腺癌相关的行
            for col in cancer_columns:
                filter_condition = filter_condition & ~cancer_data[col].astype(str).str.contains("前列腺癌", na=False)

        return cancer_data[filter_condition].reset_index(drop=True)

    # 处理二维列表格式（原有表格数据）
    if not isinstance(cancer_data, list):
        return cancer_data

    filtered_data = []
    for row in cancer_data:
        # 跳过表头行（直接保留）
        if isinstance(row, list):
            # 识别表头（编码/癌种等关键字）
            if row and (row[0] == "编码" or row[0] == "癌种"):
                filtered_data.append(row)
                continue

        # 提取癌种名称（适配不同数据格式的癌种列）
        cancer_name = ""
        if len(row) >= 2:
            cancer_name = row[1]  # 对应基础癌种数据的「患癌风险组织」列
        elif len(row) >= 1:
            cancer_name = row[0]  # 对应检测说明表格的「癌种」列

        # 性别过滤逻辑
        if gender == "男" and ("乳腺癌" in cancer_name or "乳腺导管腺癌" in cancer_name):
            continue  # 男性：跳过乳腺癌相关数据
        if gender == "女" and "前列腺癌" in cancer_name:
            continue  # 女性：跳过前列腺癌相关数据

        filtered_data.append(row)

    return filtered_data


# =========================
# 封面背景图处理函数（读取images文件夹中的封面图，优化缩放）
# =========================
def get_cover_background_image(image_dir):
    """
    从当前工作目录的images文件夹读取第一张有效图片作为封面背景
    支持常见图片格式：jpg, jpeg, png, bmp, gif
    """
    # 支持的图片格式（小写匹配，提高兼容性）
    supported_formats = (".jpg", ".jpeg", ".png", ".bmp", ".gif")

    # 遍历images目录，筛选所有有效图片
    cover_images = []
    if image_dir.exists():
        for file in os.listdir(image_dir):
            file_path = image_dir / file
            if file_path.is_file() and file.lower().endswith(supported_formats):
                cover_images.append(file_path)

    # 无有效图片时返回None，给出友好提示
    if not cover_images:
        print(f"⚠️  images文件夹中未找到有效封面图片（支持格式：{supported_formats}），将不显示封面背景图")
        return None

    # 取第一张图片作为封面背景
    cover_image_path = cover_images[0]
    print(f"✅ 已从images文件夹获取封面背景图：{cover_image_path.name}")

    # 处理封面图片：**缩小缩放比例，仅占用页面50%高度**，给文字留出更多空间
    img = Image(str(cover_image_path))
    img_width, img_height = img.drawWidth, img.drawHeight

    # 计算缩放比例（适配页面内容宽度，高度仅占用50%，大幅减少图片高度）
    content_width = PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN
    content_height = PAGE_HEIGHT - TOP_MARGIN - BOTTOM_MARGIN
    width_ratio = content_width / img_width
    height_ratio = (content_height * 0.5) / img_height  # 仅占用50%高度，给文字留空间
    scale_ratio = min(width_ratio, height_ratio, 1)  # 不放大图片，仅缩小适配

    # 应用缩放，保证图片清晰且不超出页面
    img.drawWidth = img_width * scale_ratio
    img.drawHeight = img_height * scale_ratio
    img.hAlign = "CENTER"  # 水平居中
    img.vAlign = "TOP"  # 图片靠上显示，文字在图片下方，避免重叠

    return img


# =========================
# 读取Excel文件（top20癌种特异位点v2.xlsx）并转换为PDF表格数据（新增性别过滤）
# =========================
def read_cancer_excel_to_table_data(image_dir, styles, gender):
    """
    从images文件夹读取top20癌种特异位点v2.xlsx，解析为PDF表格数据并按性别过滤
    :param image_dir: 图片/Excel目录（BASE_DIR / "images"）
    :param styles: 自定义样式对象
    :param gender: 性别（男/女），用于过滤癌种数据
    :return: reportlab Table对象（格式化后的癌种表格）
    """
    # 定义Excel文件路径（修正：使用已定义的TABLE_DIR和excel_path）
    global excel_path
    if not excel_path.exists():
        error_msg = f"❌ Excel文件不存在：{excel_path}，将显示默认提示信息"
        print(error_msg)
        # 返回包含错误提示的简易表格
        error_data = [["提示信息"], [error_msg]]
        error_table = Table(error_data, colWidths=[CONTENT_WIDTH * 0.9])
        error_table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.white),
            ("PADDING", (0, 0), (-1, -1), 6),
            ("BACKGROUND", (0, 0), (-1, 0), purple()),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("BACKGROUND", (0, 1), (-1, -1), light_purple()),
            ("TEXTCOLOR", (0, 1), (-1, -1), colors.white),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTNAME", (0, 0), (-1, 0), FONT),
            ("FONTSIZE", (0, 0), (-1, 0), 11),
            ("FONTNAME", (0, 1), (-1, -1), FONT),
            ("FONTSIZE", (0, 1), (-1, -1), 10),
        ]))
        return error_table, None

    try:
        # 读取Excel文件（支持多sheet，默认读取第一个sheet；自动识别表头，指定openpyxl引擎）
        excel_df = pd.read_excel(excel_path, sheet_name=0, header=0, engine="openpyxl")
        print(f"✅ 成功读取Excel文件：{excel_path.name}，共{len(excel_df)}行{len(excel_df.columns)}列数据")

        # 核心新增：按性别过滤Excel数据（移除对应癌种）
        excel_df = filter_cancer_by_gender(excel_df, gender)
        print(f"✅ 已按{gender}性过滤Excel数据，剩余{len(excel_df)}行有效数据")

        # 数据清洗：去除空行、空列，重置索引
        excel_df = excel_df.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
        # 替换单元格空值为空白字符串
        excel_df = excel_df.fillna("")

        # 转换为PDF表格所需的二维列表格式（表头+内容）
        table_header = [str(col) for col in excel_df.columns]
        table_content = excel_df.values.tolist()
        # 拼接完整表格数据（表头+内容）
        table_data = [table_header] + table_content

        # 定义列宽度（自适应分配，根据列数均分页面宽度的90%）
        col_count = len(table_header)
        col_widths = [(CONTENT_WIDTH * 0.9) / col_count for _ in range(col_count)]

        # 使用已有的自动换行表格函数创建格式化表格
        excel_table = create_wrapped_table(
            data=table_data,
            col_widths=col_widths,
            style=styles["TableCell"],
            risk_col_idx=None,  # Excel表格无需风险着色
            risk_image_cols=None  # Excel表格无风险图片
        )

        # 设置Excel表格专属样式
        # 设置Excel表格专属样式
        excel_table_style = TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.white),
            ("PADDING", (0, 0), (-1, -1), 6),
            ("BACKGROUND", (0, 0), (-1, 0), purple()),
            ("BACKGROUND", (0, 1), (-1, -1), light_purple()),
            ("FONTNAME", (0, 0), (-1, 0), FONT),
            ("FONTSIZE", (0, 0), (-1, 0), 10.5),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTWEIGHT", (0, 0), (-1, 0), "bold"),
            ("FONTNAME", (0, 1), (-1, -1), FONT),
            ("FONTSIZE", (0, 1), (-1, -1), 9),
            ("TEXTCOLOR", (0, 1), (-1, -1), colors.white),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ])
        excel_table.setStyle(excel_table_style)

        return excel_table, excel_df

    except Exception as e:
        error_msg = f"❌ 解析Excel文件失败：{str(e)}"
        print(error_msg)
        # 返回错误提示表格
        error_data = [["提示信息"], [error_msg]]
        error_table = Table(error_data, colWidths=[CONTENT_WIDTH * 0.9])
        error_table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.white),
            ("PADDING", (0, 0), (-1, -1), 6),
            ("BACKGROUND", (0, 0), (-1, 0), purple()),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("BACKGROUND", (0, 1), (-1, -1), light_purple()),
            ("TEXTCOLOR", (0, 1), (-1, -1), colors.white),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTNAME", (0, 0), (-1, 0), FONT),
            ("FONTSIZE", (0, 0), (-1, 0), 11),
            ("FONTNAME", (0, 1), (-1, -1), FONT),
            ("FONTSIZE", (0, 1), (-1, -1), 10),
        ]))
        return error_table, None


# =========================
# 从TSV数据中统计癌种特异位点数量（无名称映射，直接按原始癌种分组）
# =========================
def count_cancer_specific_sites(variants_df):
    """
    从新格式TSV中按癌种统计突变位点数量（去重）
    """
    # 数据清洗
    clean_df = variants_df.copy()
    clean_df = clean_df[
        (clean_df["hgvs"].notna()) &
        (clean_df["hgvs"].str.strip() != "")
    ]
    
    # 按癌种分组统计唯一HGVS数量
    cancer_site_stat = clean_df.groupby("cancer")["hgvs"].nunique().to_dict()
    
    return cancer_site_stat
# =========================
# 新增：Bash脚本调用工具函数（执行外部脚本获取单个位点的AF值）
# =========================
import subprocess
def call_bash_get_af(hgvs_site, cosmic_tsv_path, bash_script_path="get_af.sh"):
    try:
        cmd = [bash_script_path, hgvs_site, str(cosmic_tsv_path)]
        result = subprocess.run(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            encoding="utf-8", timeout=10, check=True
        )
        af_str = result.stdout.strip().split()
        if len(af_str) != 3:
            print(f"⚠️  脚本返回值格式错误，位点[{hgvs_site}] → 输出[{af_str}]")
            return {"total_sample": 0, "positive_sample": 0, "frequency": 0.0}
        # 解析脚本返回的3个值：总样本、阳性样本、频率
        total_sample = int(af_str[0])
        positive_sample = int(af_str[1])
        frequency = float(af_str[2])
        print(f"✅ 脚本匹配位点[{hgvs_site}] → 总样本{total_sample} | 阳性{positive_sample} | 频率{frequency:.6f}")
        return {
            "total_sample": total_sample,
            "positive_sample": positive_sample,
            "frequency": frequency
        }
    except Exception as e:
        print(f"⚠️  调用脚本失败，位点[{hgvs_site}] → {str(e)}")
        return {"total_sample": 0, "positive_sample": 0, "frequency": 0.0}

# =========================
# 新增：正常人样本库Bash脚本调用工具函数（执行外部脚本获取单个位点的正常人群AF值）
# =========================
def call_bash_get_normal_af(hgvs_site, normal_tsv_path, bash_script_path="./get_normal_af.sh"):
    try:
        cmd = [bash_script_path, hgvs_site, str(normal_tsv_path)]
        result = subprocess.run(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            encoding="utf-8", timeout=10, check=True
        )
        af_str = result.stdout.strip().split()
        if len(af_str) != 3:
            print(f"⚠️  正常人样本库脚本返回值格式错误，位点[{hgvs_site}] → 输出[{af_str}]")
            return {"ac": 0, "an": 0, "af": 0.0}
        # 解析脚本返回的3个值：第3列AC(阳性样本数)、第4列AN(总样本数)、第5列AF(突变比例)
        ac = int(af_str[0])
        an = int(af_str[1])
        af = float(af_str[2])
        print(f"✅ 正常人样本库匹配位点[{hgvs_site}] → 总样本{an} | 阳性{ac} | 突变比例{af:.8f}")
        return {
            "ac": ac,       # 正常人中检测到该突变的样本数（TSV第3列）
            "an": an,       # 正常人样本库中该癌种总样本数（TSV第4列）
            "af": af        # 正常人中该突变的比例（TSV第5列）
        }
    except Exception as e:
        print(f"⚠️  调用正常人样本库脚本失败，位点[{hgvs_site}] → {str(e)}")
        return {"ac": 0, "an": 0, "af": 0.0}

# =========================
# 重写：加载COSMIC突变频率映射关系（替换为调用Bash脚本）
# 入参/返回值格式完全不变，原有调用逻辑无需修改
# =========================


def load_cosmic_frequency_map(patient_name, cosmic_tsv_path):
    """
    加载COSMIC频率数据（使用全局269个位点）
    """
    try:
        # 读取全局269个位点
        hgvs_cancer_df = pd.read_csv(
            GLOBAL_HGVS_CANCER_MAP_PATH,
            sep="\t",
            encoding="utf-8",
            header=0,
            names=["hgvs", "cancer"]
        )
        hgvs_list = hgvs_cancer_df["hgvs"].astype(str).str.strip().unique().tolist()
        
        # 读取COSMIC频率TSV
        cosmic_df = pd.read_csv(
            cosmic_tsv_path,
            sep="\t",
            encoding="utf-8",
            header=0
        )
        # 构建频率映射
        freq_map = {}
        for _, row in cosmic_df.iterrows():
            hgvs = str(row.get("hgvs", "")).strip()
            if hgvs and hgvs in hgvs_list:
                freq_map[hgvs] = {
                    "total_sample": int(row.get("total_sample", 0)),
                    "positive_sample": int(row.get("positive_sample", 0)),
                    "frequency": float(row.get("frequency", 0.0))
                }
        
        print(f"✅ 加载{patient_name}的COSMIC频率数据：共{len(freq_map)}个位点匹配成功")
        return freq_map
    except Exception as e:
        print(f"⚠️  加载{patient_name}的COSMIC频率数据失败：{e}")
        # 返回空映射，避免程序崩溃
        return {}
# =========================
# 生成突变癌种筛查结果分述内容
# =========================

# 修改generate_cancer_screening_narration函数的入参，移除内部的load_cosmic_frequency_map调用
def generate_cancer_screening_narration(elements, styles, cancer_site_stat, excel_df, gender, patient_name, cosmic_tsv_path, hgvs_freq_map):
    """
    生成突变癌种筛查结果分述内容并添加到PDF元素列表
    :param hgvs_freq_map: 提前加载的COSMIC频率映射表
    """
    # 添加分述章节标题
    elements.append(Paragraph("四、突变癌种筛查结果分述", styles["Header2"]))
    
    # ========== 核心修复：读取全局269个位点的映射文件（替代旧的input_df） ==========
    try:
        # 读取全局HGVS-癌种映射文件
        hgvs_cancer_df = pd.read_csv(
            GLOBAL_HGVS_CANCER_MAP_PATH,
            sep="\t",
            encoding="utf-8",
            header=0,
            names=["hgvs", "cancer"]
        )
        # 数据清洗
        hgvs_cancer_df["hgvs"] = hgvs_cancer_df["hgvs"].astype(str).str.strip().replace("nan", "")
        cancer_hgvs_list = hgvs_cancer_df[hgvs_cancer_df["hgvs"] != ""]["hgvs"].tolist()
        print(f"✅ 加载{patient_name}的全局位点数据：共{len(cancer_hgvs_list)}个有效位点")
    except Exception as e:
        print(f"⚠️  加载{patient_name}的全局位点数据失败：{e}")
        # 兜底：使用默认位点列表
        cancer_hgvs_list = ["10:g.100354449G>C", "10:g.101134235G>T", "10:g.101770343G>C"]
    
    # 筛选出有突变的癌种
    mutated_cancers = {k: v for k, v in cancer_site_stat.items() if v > 0}

    if not mutated_cancers:
        # 无突变癌种时的提示
        elements.append(Paragraph(
            "本次检测未发现任何癌种存在特异性突变位点，受检者当前癌化风险较低，建议保持健康生活方式并定期常规体检。",
            styles["Body"]
        ))
        return

    # 遍历有突变的癌种，生成详细分述
    for idx, (cancer_name, site_count) in enumerate(mutated_cancers.items(), start=1):
        # 癌种小标题
        elements.append(Paragraph(f"{idx}. {cancer_name}", styles["Header2"]))

        # 1. 基本突变信息
        elements.append(Paragraph(
            f"<b>突变位点统计：</b>检测到{cancer_name}相关特异性突变位点{site_count}个。",
            styles["Body"]
        ))
        
        # 遍历当前癌种的所有位点
        if cancer_hgvs_list:
            for hgvs in cancer_hgvs_list:
                if not hgvs or str(hgvs).strip() == "":
                    continue
                # 1. 使用提前加载的位点信息
                site_info = hgvs_freq_map.get(hgvs, {"total_sample": 0, "positive_sample": 0, "frequency": 0.0})
                total_sample = site_info["total_sample"]
                positive_sample = site_info["positive_sample"]
                frequency = site_info["frequency"]
                # 拼接癌种库句式（原有）
                cancer_freq_text = f"在COSMIC数据库中，{cancer_name}病人样本有{total_sample}个，其中{positive_sample}份样本检测到{hgvs}突变，{hgvs}突变频率为{frequency:.6f}。"
                elements.append(Paragraph(f"<b>COSMIC癌种数据库频率：</b>{cancer_freq_text}", styles["Body"]))

                # 2. 调用正常人样本库脚本，获取位点信息（已加空值校验）
                normal_site_info = call_bash_get_normal_af(
                    hgvs_site=hgvs,
                    normal_tsv_path=NORMAL_COSMIC_TSV_PATH,
                    bash_script_path="./get_normal_af.sh"
                )
                normal_an = normal_site_info["an"]  # 正常人总样本数（TSV第4列）
                normal_ac = normal_site_info["ac"]  # 正常人阳性样本数（TSV第3列）
                normal_af = normal_site_info["af"]  # 正常人突变比例（TSV第5列）
                
                # 核心修改：添加ac为0的条件判断
                if normal_ac == 0:
                    # ac为0时打印指定提示
                    normal_freq_text = f"该突变在正常人样本库中不存在。"
                else:
                    # ac不为0时使用原有句式
                    normal_freq_text = f"在正常人样本库中，在{normal_an}条染色体中有{normal_ac}条检测到{hgvs}突变，突变频率为{normal_af:.8f}。"
                elements.append(Paragraph(f"<b>正常人样本库频率：</b>{normal_freq_text}", styles["Body"]))
        
        # 2. 从Excel中获取该癌种的详细信息
        excel_info = ""
        if excel_df is not None and not excel_df.empty:
            # 查找Excel中匹配的癌种行
            cancer_columns = [col for col in excel_df.columns if
                              any(keyword in str(col) for keyword in ["癌种", "癌症", "肿瘤", "组织"])]
            if cancer_columns:
                cancer_col = cancer_columns[0]
                # 模糊匹配癌种名称
                mask = excel_df[cancer_col].astype(str).str.contains(cancer_name, na=False)
                if mask.any():
                    cancer_row = excel_df[mask].iloc[0]

                    # 提取检测基因数和位点总数
                    gene_count = ""
                    total_sites = ""
                    # 适配新列名
                    for col in excel_df.columns:
                        col_name = str(col).lower()
                        # 匹配新表格的「特异位点-基因数」列
                        if any(key in col_name for key in ["特异位点", "gene count"]):
                            gene_count = cancer_row[col]
                        # 匹配新表格的「特异位点-位点数」列
                        elif any(key in col_name for key in ["位点数", "site count"]):
                            total_sites = cancer_row[col]

                    if gene_count and total_sites:
                        mutation_ratio = site_count / int(total_sites) if str(total_sites).isdigit() else 0
                        excel_info = f"本次检测覆盖{cancer_name}相关基因{gene_count}个，总检测位点{total_sites}个，突变比例约{mutation_ratio:.6f}。"
                    elif gene_count:
                        excel_info = f"本次检测覆盖{cancer_name}相关基因{gene_count}个。"
                    elif total_sites:
                        excel_info = f"本次检测覆盖{cancer_name}相关位点{total_sites}个。"

        if excel_info:
            elements.append(Paragraph(excel_info, styles["Body"]))

        # 3. 风险等级判定
        if site_count >= 10:
            risk_desc = "高风险"
            risk_suggest = "该突变数量提示癌化风险较高，建议立即前往肿瘤科进行专项检查（如影像学检查、病理活检等），并制定个性化干预方案。"
        elif site_count >= 3:
            risk_desc = "中风险"
            risk_suggest = "该突变数量提示癌化风险中等，建议在1-3个月内完成对应癌种的专项筛查，并密切关注身体相关症状。"
        else:
            risk_desc = "低风险"
            risk_suggest = "该突变数量提示癌化风险较低，但仍需重视，建议6个月内复查，并调整生活方式降低癌变风险。"

        elements.append(Paragraph(f"<b>风险等级：</b>{risk_desc}", styles["Body"]))
        elements.append(Paragraph(f"<b>医学建议：</b>{risk_suggest}", styles["Body"]))

        # 4. 针对性干预建议（按癌种适配）
        elements.append(Paragraph("<b>针对性干预建议：</b>", styles["Body"]))
        
        # 不同癌种的专属建议
        cancer_advices = {
            "肺癌": [
                "立即戒烟，包括电子烟，远离烟草相关产品",
                "定期进行肺功能检测和胸部CT检查（首选低剂量螺旋CT）",
                "避免接触石棉、氡气等致癌物质，工作环境需做好防护",
                "坚持有氧运动（如快走、游泳），改善肺功能"
            ],
            "乳腺癌": [
                "每月进行乳腺自我检查，每年做一次乳腺超声/钼靶检查",
                "控制雌激素暴露，避免长期使用含雌激素的保健品",
                "保持健康体重，避免高脂饮食",
                "母乳喂养（如有条件），降低乳腺癌风险"
            ],
            "肝癌": [
                "戒酒，避免饮用含酒精饮料",
                "接种乙肝疫苗，定期检查肝功能和乙肝五项",
                "避免食用发霉的谷物、坚果（黄曲霉素）",
                "控制脂肪肝，规律运动，低脂饮食"
            ],
            "大肠癌": [
                "50岁以上每年做一次肠镜检查",
                "增加膳食纤维摄入，减少红肉和加工肉类食用",
                "保持规律排便，避免便秘",
                "控制体重，每周至少150分钟中等强度运动"
            ],
            # 可添加更多癌种的建议
        }
        
        # 获取当前癌种的建议，无则用通用建议
        specific_advices = cancer_advices.get(cancer_name, [
            "戒烟限酒，保持规律作息",
            "均衡饮食，增加新鲜蔬果摄入",
            "定期进行对应癌种的专项筛查",
            "保持积极心态，适度运动"
        ])

        # 添加建议列表
        for adv_idx, advice in enumerate(specific_advices, 1):
            elements.append(Paragraph(f"{adv_idx}. {advice}", styles["Body"]))

        # 癌种之间添加间隔
        elements.append(Spacer(1, 10))

    # 添加总体建议
    elements.append(Paragraph("<b>总体建议：</b>", styles["Header2"]))
    elements.append(Paragraph(
        f"本次检测共发现{len(mutated_cancers)}种癌种存在特异性突变位点，其中高风险{len([k for k, v in mutated_cancers.items() if v >= 10])}种，"
        f"中风险{len([k for k, v in mutated_cancers.items() if 3 <= v < 10])}种，低风险{len([k for k, v in mutated_cancers.items() if v < 3])}种。"
        "建议优先处理高/中风险癌种的筛查和干预，定期复查并遵循临床医生的专业指导。",
        styles["Body"]
    ))
# =========================
# 基础配置（路径+字体+视觉主题）
# =========================
BASE_DIR = Path(__file__).resolve().parent
# 字体配置 - 修改为宋体
FONT_DIR = BASE_DIR / "fonts"
FONT_DIR.mkdir(exist_ok=True)
# 宋体字体文件路径（确保simsun.ttf存在于fonts目录）
FONT_PATH = str(FONT_DIR / "simsun.ttc")  # 宋体文件名通常是simsun.ttf/SimSun.ttf

# 兼容处理：如果字体文件不存在，使用默认字体
if os.path.exists(FONT_PATH):
    pdfmetrics.registerFont(TTFont("SimSun", FONT_PATH))  # 注册宋体，字体名称设为"SimSun"
    FONT = "SimSun"  # 全局字体变量改为宋体
else:
    # 备选方案：尝试系统内置宋体（部分环境支持）
    try:
        pdfmetrics.registerFont(TTFont("SimSun", "SimSun"))
        FONT = "SimSun"
        print(f"⚠️  未找到本地宋体文件，使用系统内置宋体")
    except:
        FONT = "Helvetica"
        print(f"⚠️  宋体字体文件不存在：{FONT_PATH}，使用默认字体")

# 输出目录
OUTPUT_DIR = BASE_DIR/ "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# 图片配置（当前工作目录下的images文件夹，封面图/Excel存放于此）
IMAGE_DIR = BASE_DIR / "images"
# 1. 先定义table目录路径
TABLE_DIR = BASE_DIR / "table"  # 注意：目录名用字符串包裹，不要省略引号
TABLE_DIR.mkdir(exist_ok=True)  # 自动创建table目录，避免文件不存在报错
# 2. 再拼接Excel文件名
excel_path = TABLE_DIR /"癌症统计表.xlsx"
IMAGE_DIR.mkdir(exist_ok=True)
TARGET_IMAGE_PATH = IMAGE_DIR / "cancer_gene_chart.jpeg"  # 保留路径，仅作兼容
LOGO_PATH = IMAGE_DIR / "logo.png"  # 可选公司LOGO
# 封面背景图：与其他图片/Excel同目录（images文件夹）
TMB_IMAGE_PATH = IMAGE_DIR / "lung_carcinoma_unique_sites_site_sample_distribution.png"  # TBM.png存放在images文件夹
TMB2_IMAGE_PATH = IMAGE_DIR / "liver_unique_sites_site_sample_distribution.png"  # 新增：TMB2图片路径

# 临时目录（存储风险条图片）
TEMP_DIR = BASE_DIR / "temp"
TEMP_DIR.mkdir(exist_ok=True)

# TSV文件路径模板
TSV_ROOT_DIR = BASE_DIR
# 新增：癌症形成机制图片路径配置
CANCER_FORMATION_IMAGE_PATH = IMAGE_DIR / "癌症成因.png"  # 从images文件夹读取癌症成因图
# 补充缺失的PERSON_INFO_TEMPLATE定义（修复原有代码潜在报错）
PERSON_INFO_TEMPLATE = "{patient_name}input.tsv"
# ========== 核心配置：全局HGVS-癌种映射文件（269个位点） ==========
GLOBAL_HGVS_CANCER_MAP_PATH = TABLE_DIR / "variant_cancer_map_cn.tsv"  # 269个位点的映射文件
# 废弃的旧患者位点文件模板（标记为废弃）
OLD_VARIANT_TSV_TEMPLATE = "{patient_name}input-位点.tsv"
#正常人
NORMAL_COSMIC_TSV_PATH = TABLE_DIR / "cosmic_specific_AC_AN_AF_final.tsv"

# 页面尺寸与边距
PAGE_WIDTH, PAGE_HEIGHT = A4
LEFT_MARGIN = 35
RIGHT_MARGIN = 35
TOP_MARGIN = 30
BOTTOM_MARGIN = 30
CONTENT_WIDTH = PAGE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN
CONTENT_HEIGHT = PAGE_HEIGHT - TOP_MARGIN - BOTTOM_MARGIN

# 替换为新配色（蓝白+浅灰+淡黄绿，低饱和度）
# 替换为新配色（蓝白+浅灰+淡黄绿，低饱和度）
THEME_COLOR = {
    "primary": colors.Color(0.15, 0.35, 0.65, 1),  # 品牌蓝（沉稳不刺眼）
    "secondary": colors.Color(0.92, 0.94, 0.96, 1),  # 浅灰（替代浅蓝灰，更舒缓）
    "accent": colors.Color(0.95, 0.98, 0.85, 1),  # 淡黄绿（替代亮黄，仅提示不刺眼）
    "text": colors.Color(0.1, 0.1, 0.1, 1),  # 深灰（比原配色更浅，减少视觉冲击）
    "risk_high": colors.Color(0.7, 0.2, 0.2, 1),  # 暗红（降低饱和度，避免恐慌）
    "risk_mid": colors.Color(0.8, 0.7, 0.2, 1),  # 淡黄（替代亮橙，更柔和）
    "risk_low": colors.Color(0.2, 0.6, 0.3, 1),  # 淡绿（降低饱和度）
    "border_gray": colors.Color(0.85, 0.87, 0.89, 1),  # 新增：表格边框浅灰
    "header_yellow": colors.Color(1.0, 0.96, 0.79, 1)  # 新增：淡黄色表头背景（#FFEDC7）
}


# =========================
# 核心优化1：文字自动换行工具
# =========================
def wrap_text(text, max_width, style):
    """自动拆分长文本，避免表格文字溢出（兼容空文本）"""
    if not text or str(text).strip() == "":
        return [""]
    words = str(text).split()
    lines = []
    current_line = []
    current_width = 0

    for word in words:
        word_width = pdfmetrics.stringWidth(word + " ", style.fontName, style.fontSize)
        if current_width + word_width > max_width:
            if current_line:
                lines.append(" ".join(current_line))
                current_line = [word]
                current_width = word_width
            else:
                lines.append(word)
                current_width = 0
        else:
            current_line.append(word)
            current_width += word_width
    if current_line:
        lines.append(" ".join(current_line))
    return lines


def create_wrapped_table(data, col_widths, style, risk_col_idx=None, risk_image_cols=None):
    """
    创建支持文字换行+风险着色+风险图片的表格（无名称映射相关逻辑）
    :param risk_image_cols: 存储风险条图片的列索引列表
    """
    # 安全校验：确保所有行的列数与列宽度列表长度一致
    row_cols = len(col_widths)
    for row_idx, row in enumerate(data):
        if len(row) != row_cols:
            # 补全缺失的列或截断多余的列，避免索引错误
            if len(row) < row_cols:
                row.extend([""] * (row_cols - len(row)))
            else:
                data[row_idx] = row[:row_cols]

    # 处理表头
    header_style = ParagraphStyle(
        name="TableHeader_Wrapped",
        parent=style,
        bold=True,
        fontSize=style.fontSize + 0.5
    )
    wrapped_data = [[Paragraph(cell, header_style) for cell in data[0]]]

    # 处理内容行（自动换行+风险着色+风险图片）
    for row_idx, row in enumerate(data[1:]):
        wrapped_row = []
        for idx, cell in enumerate(row):
            # 安全校验：避免idx超出col_widths范围
            if idx >= len(col_widths):
                cell = ""

            # 如果是风险图片列，直接添加图片对象
            if risk_image_cols and idx in risk_image_cols:
                if isinstance(cell, Image):
                    wrapped_row.append(cell)
                else:
                    wrapped_row.append(Paragraph("", style))
                continue

            # 普通列处理自动换行
            max_width = col_widths[idx] - 10 if idx < len(col_widths) else 50
            wrapped_lines = wrap_text(cell, max_width, style)
            wrapped_text = "<br/>".join(wrapped_lines)  # 修复：原代码为"<br"，缺少/>导致换行失效

            # 根据风险等级设置单元格样式（仅对非图片列生效）
            if risk_col_idx is not None and idx == risk_col_idx and not (risk_image_cols and idx in risk_image_cols):
                risk_level = str(cell).strip()
                if "高" in risk_level:
                    cell_style = ParagraphStyle(
                        name=f"RiskHigh_Cell_{row_idx}_{idx}",
                        parent=style,
                        textColor=colors.white,
                        backColor=THEME_COLOR["risk_high"]
                    )
                elif "中" in risk_level:
                    cell_style = ParagraphStyle(
                        name=f"RiskMid_Cell_{row_idx}_{idx}",
                        parent=style,
                        textColor=colors.white,
                        backColor=THEME_COLOR["risk_mid"]
                    )
                elif "低" in risk_level or "阴" in risk_level:
                    cell_style = ParagraphStyle(
                        name=f"RiskLow_Cell_{row_idx}_{idx}",
                        parent=style,
                        textColor=colors.white,
                        backColor=THEME_COLOR["risk_low"]
                    )
                else:
                    cell_style = style
            else:
                cell_style = style

            wrapped_row.append(Paragraph(wrapped_text, cell_style))
        wrapped_data.append(wrapped_row)

    # 创建表格
    table = Table(wrapped_data, colWidths=col_widths)
    return table


# =========================
# 核心优化2：样式定义（调整封面标题间距，减少垂直占用）
# =========================
def get_custom_styles():
    styles = getSampleStyleSheet()

    # 封面标题（**缩小spaceAfter，减少垂直间距**）
    styles.add(ParagraphStyle(
        name="CoverTitle",
        fontName=FONT,
        fontSize=22,
        alignment=1,
        spaceAfter=15,  # 从25改为15，减少标题间的垂直间距
        textColor=THEME_COLOR["primary"],
        bold=True
    ))

    # 报告副标题（**缩小spaceAfter，减少垂直间距**）
    styles.add(ParagraphStyle(
        name="CoverSubtitle",
        fontName=FONT,
        fontSize=14,
        alignment=1,
        spaceAfter=20,  # 从40改为20，减少副标题后的垂直间距
        textColor=THEME_COLOR["text"]
    ))

    # 一级标题
    styles.add(ParagraphStyle(
        name="Header1",
        fontName=FONT,
        fontSize=16,
        spaceBefore=25,
        spaceAfter=12,
        textColor=THEME_COLOR["primary"],
        bold=True,
    ))

    # 二级标题（小节标题）
    styles.add(ParagraphStyle(
        name="Header2",
        fontName=FONT,
        fontSize=14,
        spaceBefore=18,
        spaceAfter=8,
        textColor=THEME_COLOR["text"],
        bold=True,
        leftIndent=5
    ))
    
    # 三级标题（小节子标题）
    styles.add(ParagraphStyle(
    name="Header3",
    fontName=FONT,
    fontSize=12,
    spaceBefore=15,
    spaceAfter=6,
    textColor=THEME_COLOR["text"],
    bold=True,
    leftIndent=10
))

    # 正文样式
    styles.add(ParagraphStyle(
        name="Body",
        fontName=FONT,
        fontSize=11,
        leading=19,
        spaceAfter=10,
        textColor=THEME_COLOR["text"],
        leftIndent=0
    ))

    # 表格内容样式
    styles.add(ParagraphStyle(
        name="TableCell",
        fontName=FONT,
        fontSize=10,
        leading=12,
        textColor=THEME_COLOR["text"],
        alignment=1  # 居中对齐
    ))

    # 表格表头样式
    styles.add(ParagraphStyle(
        name="TableHeader",
        fontName=FONT,
        fontSize=10.5,
        leading=16,
        textColor=colors.white,
        alignment=1,
        bold=True
    ))

    # 风险提示样式
    styles.add(ParagraphStyle(
        name="RiskNote",
        fontName=FONT,
        fontSize=10,
        leading=16,
        textColor=THEME_COLOR["risk_high"],
        bold=True,
        spaceAfter=8
    ))

    return styles


# =========================
# 核心优化3：表格样式模板
# =========================
def get_table_style(template_type="default"):
    """预定义表格样式模板"""
    base_style = [
        ("GRID", (0, 0), (-1, -1), 0.5, colors.white),
        ("PADDING", (0, 0), (-1, -1), 6),
        ("BACKGROUND", (0, 0), (-1, 0), purple()),
        ("BACKGROUND", (0, 1), (-1, -1), light_purple()),
        ("FONTNAME", (0, 0), (-1, 0), FONT),
        ("FONTSIZE", (0, 0), (-1, 0), 10.5),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTWEIGHT", (0, 0), (-1, 0), "bold"),
        ("FONTNAME", (0, 1), (-1, -1), FONT),
        ("FONTSIZE", (0, 1), (-1, -1), 10),
        ("TEXTCOLOR", (0, 1), (-1, -1), colors.white),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]

    if template_type == "message":
        base_style.extend([
            ("BACKGROUND", (0, 0), (-1, 0), purple()),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("BACKGROUND", (0, 1), (-1, -1), light_purple()),
            ("TEXTCOLOR", (0, 1), (-1, -1), colors.white),
        ])

    return TableStyle(base_style)


# =========================
# 图片插入工具（保留，用于其他图片场景）
# =========================
def insert_image_with_fit(image_path, max_width=None, max_height=None):
    """插入图片并自动适配页面（避免超出）"""
    if not os.path.exists(image_path):
        return Paragraph("【图片缺失】", ParagraphStyle(
            name="ImageMissingNote",
            fontName=FONT,
            fontSize=10,
            textColor=THEME_COLOR["risk_high"],
            alignment=1
        ))

    img = Image(str(image_path))
    default_max_width = CONTENT_WIDTH * 0.9
    default_max_height = CONTENT_HEIGHT * 0.4
    max_width = max_width or default_max_width
    max_height = max_height or default_max_height

    # 修复：使用正确的Image对象属性
    img_width, img_height = img.drawWidth, img.drawHeight
    width_ratio = max_width / img_width
    height_ratio = max_height / img_height
    scale_ratio = min(width_ratio, height_ratio, 1)

    img.drawWidth = img_width * scale_ratio
    img.drawHeight = img_height * scale_ratio
    img.hAlign = "CENTER"
    return img

def add_chapter_cover(elements, chapter_title, image_name, styles):
    """
    为章节添加独立封面页（图片+标题单独占一页）
    :param elements: PDF元素列表
    :param chapter_title: 章节标题（如"第一章节：检测基础信息"）
    :param image_name: images文件夹中的图片文件名（如"章节1封面.png"）
    :param styles: 样式对象
    """
#    # 强制插入分页符，新建空白页作为章节封面
    elements.append(PageBreak())
    elements.append(NextPageTemplate("Body"))  # 沿用正文页面模板
    
    # 插入章节封面图（从images文件夹读取）
    cover_img_path = IMAGE_DIR / image_name
    if os.path.exists(cover_img_path):
        # 图片自适应页面（占页面80%宽度，居中显示）
        cover_img = insert_image_with_fit(
            str(cover_img_path),
            max_width=CONTENT_WIDTH * 0.8,
            max_height=CONTENT_HEIGHT * 0.6
        )
        elements.append(cover_img)
        elements.append(Spacer(1, 30))  # 图片与标题间距
    
    # 章节标题（居中、加粗、放大）
    chapter_cover_style = ParagraphStyle(
        name="ChapterCoverTitle",
        fontName=FONT,
        fontSize=18,
        alignment=1,  # 居中
        spaceAfter=20,
        textColor=THEME_COLOR["primary"],
        bold=True
    )
    elements.append(Paragraph(chapter_title, chapter_cover_style))
    
    # 再次插入分页符，确保章节封面页结束后，下一页才显示章节正文
    elements.append(PageBreak())
# =========================
# 新增：插入PDF文件作为图片的工具函数
# =========================
from PyPDF2 import PdfReader
from reportlab.pdfgen import canvas
import io


def insert_pdf_as_image(pdf_path, max_width=None, max_height=None):
    """
    将PDF文件的第一页转为图片插入PDF报告（适配页面尺寸）
    :param pdf_path: PDF文件路径
    :param max_width: 最大宽度
    :param max_height: 最大高度
    :return: reportlab的Image对象/缺失提示
    """
    if not os.path.exists(pdf_path):
        return Paragraph("【PDF文件缺失】", ParagraphStyle(
            name="PDFMissingNote",
            fontName=FONT,
            fontSize=10,
            textColor=THEME_COLOR["risk_high"],
            alignment=1
        ))

    try:
        # 读取PDF第一页并转为临时图片（PNG）
        reader = PdfReader(pdf_path)
        page = reader.pages[0]

        # 创建临时缓冲区存储图片
        temp_buffer = io.BytesIO()
        c = canvas.Canvas(temp_buffer, pagesize=A4)

        # 渲染PDF页面到画布
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)
        c.drawImage(pdf_path, 0, 0, width=page_width, height=page_height)
        c.save()

        # 重置缓冲区指针
        temp_buffer.seek(0)

        # 创建Image对象并适配尺寸
        img = Image(temp_buffer)
        default_max_width = CONTENT_WIDTH * 0.9
        default_max_height = CONTENT_HEIGHT * 0.4
        max_width = max_width or default_max_width
        max_height = max_height or default_max_height

        # 计算缩放比例
        img_width, img_height = img.drawWidth, img.drawHeight
        width_ratio = max_width / img_width
        height_ratio = max_height / img_height
        scale_ratio = min(width_ratio, height_ratio, 1)

        # 应用缩放
        img.drawWidth = img_width * scale_ratio
        img.drawHeight = img_height * scale_ratio
        img.hAlign = "CENTER"

        return img
    except Exception as e:
        return Paragraph(f"【PDF解析失败：{str(e)}】", ParagraphStyle(
            name="PDFErrorNote",
            fontName=FONT,
            fontSize=10,
            textColor=THEME_COLOR["risk_high"],
            alignment=1
        ))
# =========================
# TSV数据读取（补充肺癌数据，癌种名称与表格完全对齐，无映射）
# =========================

# =========================
# 新增：读取HGVS-癌种映射文件
# =========================
def read_hgvs_cancer_map(map_tsv_path):
    """
    读取variant_cancer_map_cn.tsv文件，构建HGVS到癌种的映射字典
    :param map_tsv_path: 映射文件路径
    :return: dict，格式：{hgvs: 癌种}
    """
    hgvs_cancer_map = {}
    if not os.path.exists(map_tsv_path):
        print(f"❌ HGVS-癌种映射文件不存在：{map_tsv_path}")
        return hgvs_cancer_map
    
    try:
        # 读取TSV文件（无表头，两列：HGVS、癌种）
        df = pd.read_csv(
            map_tsv_path, 
            sep="\t", 
            encoding="utf-8",
            header=None,
            names=["hgvs", "cancer"]
        )
        # 去重：保留每个HGVS对应的第一个癌种
        df = df.drop_duplicates(subset=["hgvs"], keep="first")
        # 构建映射字典
        hgvs_cancer_map = dict(zip(df["hgvs"], df["cancer"]))
        print(f"✅ 读取HGVS-癌种映射文件，共{len(hgvs_cancer_map)}个唯一HGVS位点")
        return hgvs_cancer_map
    except Exception as e:
        print(f"⚠️  读取HGVS-癌种映射文件失败：{e}")
        return hgvs_cancer_map
        
        
def read_patient_info_tsv(patient_name):
    patient_dir = TSV_ROOT_DIR / patient_name
    tsv_path = patient_dir / PERSON_INFO_TEMPLATE.format(patient_name=patient_name)

    if not tsv_path.exists():
        tsv_path = TSV_ROOT_DIR / PERSON_INFO_TEMPLATE.format(patient_name=patient_name)

    try:
        df = pd.read_csv(tsv_path, sep="\t", encoding="utf-8")
        sample = df.iloc[0].to_dict()
        default_fields = {
            "phone": "", "hospital": "", "doctor": "",
            "send_date": "", "receive_date": "", "report_date": "", "purpose": ""
        }
        sample.update({k: sample.get(k, v) for k, v in default_fields.items()})

        # 新增：性别数据清洗，统一格式为「男」或「女」
        if "gender" in sample:
            gender = str(sample["gender"]).strip().lower()
            if gender in ["男", "男性", "m", "male"]:
                sample["gender"] = "男"
            elif gender in ["女", "女性", "f", "female"]:
                sample["gender"] = "女"
            else:
                sample["gender"] = "男"  # 默认值

        print(f"✅ 读取{patient_name}个人信息：{tsv_path}")
        return sample
    except Exception as e:
        print(f"⚠️  读取{patient_name}个人信息失败，使用默认数据：{e}")
        return {
            "name": patient_name, "age": "45", "gender": "男",
            "sample_id": f"SZ2024{patient_name[-1:]}01",
            "phone": "", "hospital": "深圳星辰合作医院", "doctor": "基因检测科医生",
            "send_date": "2024-01-01", "receive_date": "2024-01-02",
            "report_date": "2024-01-05", "purpose": "多系统泛癌早筛"
        }


def read_variant_tsv(patient_name):
    """
    核心修改：不再读取患者目录下的旧位点文件，改为读取全局269个位点的映射文件
    """
    # ========== 第一步：读取全局HGVS-癌种映射文件（269个位点） ==========
    if not GLOBAL_HGVS_CANCER_MAP_PATH.exists():
        raise FileNotFoundError(f"全局HGVS-癌种映射文件不存在：{GLOBAL_HGVS_CANCER_MAP_PATH}")
    
    # 读取269个位点的完整映射表
    hgvs_cancer_df = pd.read_csv(
        GLOBAL_HGVS_CANCER_MAP_PATH,
        sep="\t",
        encoding="utf-8",
        header=0,  # 映射文件有表头（HGVS, 癌种）
        names=["hgvs", "cancer"]  # 确保列名统一
    )
    
    # ========== 第二步：数据清洗（保留269个有效位点） ==========
    # 彻底清洗空值/无效值
    hgvs_cancer_df["hgvs"] = hgvs_cancer_df["hgvs"].astype(str).str.strip().replace("nan", "")
    hgvs_cancer_df = hgvs_cancer_df[
        (hgvs_cancer_df["hgvs"] != "") & 
        (hgvs_cancer_df["cancer"] != "")
    ].drop_duplicates(subset=["hgvs"]).reset_index(drop=True)
    
    # 验证是否读取到269个位点
    total_sites = len(hgvs_cancer_df)
    if total_sites != 269:
        print(f"⚠️  全局映射文件位点数量异常：读取到{total_sites}个，预期269个")
    else:
        print(f"✅ 成功读取全局HGVS-癌种映射文件：共{total_sites}个有效位点")
    
    # ========== 第三步：补充患者位点所需的其他字段 ==========
    hgvs_cancer_df["gene"] = "未知基因"  # 可根据实际需求补充基因名
    hgvs_cancer_df["risk_level"] = "高"  # 默认风险等级
    hgvs_cancer_df["specific_sites"] = 1  # 每个位点计1个
    hgvs_cancer_df["total_sites"] = 249188  # 总位点数
    
    print(f"✅ 为{patient_name}加载全局位点数据：共{len(hgvs_cancer_df)}个位点，关联{hgvs_cancer_df['cancer'].nunique()}种癌种")
    return hgvs_cancer_df
# =========================
# PDF生成主函数（核心：Excel表格性别过滤，封面内容不溢出，修复10的次方显示）
# =========================
def generate_beautiful_pdf(sample, variants, output_pdf):
    styles = get_custom_styles()
    gender = sample["gender"]

    # 新增：自定义页面模板，添加浅灰背景
    def custom_page(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(THEME_COLOR["secondary"])  # 浅灰背景
        canvas.rect(0, 0, A4[0], A4[1], fill=1, stroke=0)
        canvas.restoreState()

    # 修改原文档模板创建逻辑
    doc = BaseDocTemplate(
        str(output_pdf),
        pagesize=A4,
        leftMargin=LEFT_MARGIN,
        rightMargin=RIGHT_MARGIN,
        topMargin=TOP_MARGIN,
        bottomMargin=BOTTOM_MARGIN,
        showBoundary=False
    )
    # 应用自定义页面背景
    cover_frame = Frame(LEFT_MARGIN, BOTTOM_MARGIN, CONTENT_WIDTH, CONTENT_HEIGHT, id="coverFrame", showBoundary=False)
    body_frame = Frame(LEFT_MARGIN, BOTTOM_MARGIN, CONTENT_WIDTH, CONTENT_HEIGHT - 10, id="bodyFrame", showBoundary=False)
    doc.addPageTemplates([
        PageTemplate(id="Cover", frames=[cover_frame], onPage=custom_page),  # 封面加背景
        PageTemplate(id="Body", frames=[body_frame], onPage=custom_page)     # 正文加背景
    ])
    # 其余代码不变

    elements = []

    # 封面设计（核心调整：减少元素间距，确保所有内容在第一页）
    elements.append(NextPageTemplate("Cover"))

    # 步骤1：插入封面背景图（从当前工作目录的images文件夹读取，已优化缩放）
    cover_bg_image = get_cover_background_image(IMAGE_DIR)
    if cover_bg_image:
        elements.append(cover_bg_image)
        elements.append(Spacer(1, 15))  # 图片与LOGO之间的间距从30改为15，减少垂直占用

    # 步骤2：保留原有LOGO展示逻辑（同目录images文件夹）
    if os.path.exists(LOGO_PATH):
        logo = Image(str(LOGO_PATH))
        logo.drawWidth = 120
        logo.drawHeight = 40
        logo.hAlign = "CENTER"
        elements.append(logo)
        elements.append(Spacer(1, 10))  # LOGO与标题之间的间距从20改为10，减少垂直占用

    # 步骤3：封面标题与副标题（已在样式中缩小间距）
    elements.append(Paragraph("OncoSeek平台", styles["CoverTitle"]))
    elements.append(Paragraph("科学防癌定量精准检测报告", styles["CoverTitle"]))
    elements.append(Paragraph("CELL AND GENE TECHNOLOGIES BENEFITING HUMAN HEALTH", styles["CoverSubtitle"]))
    elements.append(Paragraph("预见健康 · 遇见你", styles["CoverSubtitle"]))

    # 步骤4：受检者信息（**缩小Spacer高度，从200改为100，大幅减少空白**）
    elements.append(Spacer(1, 100))  # 从200改为100，减少标题与信息之间的空白
    elements.append(Paragraph(f"受检者：{sample['name']}", styles["Body"]))
    elements.append(Paragraph(f"样本编号：{sample['sample_id']}", styles["Body"]))
    elements.append(Paragraph(f"报告日期：{sample['report_date']}", styles["Body"]))

    # 步骤5：PageBreak移到所有封面内容之后，确保封面内容在第一页
    elements.append(PageBreak())
    elements.append(NextPageTemplate("Body"))


        # 致客户信
    elements.append(Paragraph("致受检者", styles["Header1"]))
    gender_title = "先生" if sample["gender"] == "男" else "女士"
    elements.append(Paragraph(f"尊敬的{sample['name']}{gender_title}：", styles["Body"]))
    elements.append(Paragraph(
        "首先，感谢您对我们的信任与选择。至此，您的个人基因图谱已全面完成检测与系统解读。这份基因图谱记录的是属于您独有的生命遗传信息，是人体最基础、最核心的生物学信息之一，也是个人健康身份的重要组成部分。通过对这些遗传信息的科学解析，我们能够从分子层面更深入地了解个体在疾病易感性、遗传特征以及潜在健康风险等方面的生物学基础。",
        styles["Body"]
    ))
    elements.append(Paragraph("在本次检测过程中，我们基于先进的基因检测技术与严格的质量控制体系，对您的基因组数据进行了全面、严谨且多维度的分析与评估。通过对遗传性疾病相关基因位点、肿瘤易感基因以及多种已被临床研究证实具有重要意义的突变位点进行系统筛查，并结合权威遗传数据库、临床研究数据及大规模人群参考数据进行比对分析，我们完成了本次遗传性疾病风险与肿瘤相关风险的综合评估。同时，我们还根据检测结果中的突变位点信息、风险等级分层、人群流行病学数据以及国际临床数据库资料，对潜在健康风险进行了科学解释，并在此基础上形成了结构化、系统化且具有可操作性的健康管理与干预指导建议。",
       styles["Body"]
    ))
    elements.append(Paragraph("我们希望，本报告不仅能够帮助您更加科学、精准、全面地了解自身的遗传特征和潜在健康风险，还能够为您在未来的健康管理中提供重要参考依据。通过对相关风险的提前识别与合理管理，您可以在疾病预防、生活方式优化、定期健康监测以及个体化健康管理等方面做出更加理性和有效的决策，从而更好地维护自身健康水平，降低潜在疾病风险，并持续提升生活质量与长期健康状态。",
       styles["Body"]
    ))
    elements.append(Paragraph(
        "同时，我们高度重视个人隐私保护与生物信息安全。本报告的生成严格遵循相关法律法规以及个人隐私与生物信息保护原则。您的基因数据仅用于本次检测与分析过程，所有数据在存储、传输与处理过程中均采用加密技术进行严格保护，并由专业系统进行安全管理。在未获得您本人明确授权的情况下，您的任何基因信息与个人数据均不会被用于其他用途，也不会向任何第三方机构或个人泄露，以确保您的隐私权益得到充分保障。",
        styles["Body"]
    ))
    elements.append(Paragraph(
        "需要特别说明的是，本报告的分析结果基于本次送检样本所得数据，仅对该样本负责。报告中的所有内容均属于健康风险评估、健康指导及健康管理参考信息，其目的在于帮助您了解潜在健康风险，并提供科学的健康管理建议。本报告不构成临床医学诊断结论，也不能作为疾病确诊、治疗方案制定或医疗行为实施的直接依据。由于个体健康状况受到遗传因素、生活方式、环境因素及其他多种复杂因素的共同影响，因此基因检测结果应结合临床表现、医学检查以及专业医生评估进行综合判断。",
        styles["Body"]
    ))
    elements.append(Paragraph(
        "因此，请您理解，本报告结果仅用于健康风险评估与健康管理参考，不能替代专业医师的医学诊断与治疗建议。如果您需要进一步的医学判断、临床诊疗建议、专项医学检查或针对性的临床干预措施，建议您咨询具有资质的专业临床医师或医疗机构，在医生的指导下进行进一步随访、检查与个体化健康管理，以获得更加全面、科学和安全的医疗支持。",
        styles["Body"]
    ))
    
    elements.append(Spacer(1, 12))

    # 受检者信息
    # 第一章节：检测基础信息
    # 新增第一章节封面（图片名可自定义，需放在images文件夹）
    add_chapter_cover(elements, "第一章节：检测基础信息", "章节1封面.png", styles)
    elements.append(Paragraph("第一章节：检测基础信息", styles["Header1"]))
    elements.append(Paragraph("一、受检者基本信息", styles["Header2"]))
    info_data = [
        ["项目", "详情", "项目", "详情"],
        ["姓名", sample["name"], "性别", sample["gender"]],
        ["年龄", sample["age"], "联系方式", sample["phone"]],
        ["送检单位", sample["hospital"], "条码号", sample.get("barcode", "无")],
        ["送检医生", sample["doctor"], "送检样本", "外周血"],
        ["送样日期", sample["send_date"], "收样日期", sample["receive_date"]],
        ["报告日期", sample["report_date"], "样本编号", sample["sample_id"]],
        ["检测项目", "多系统泛癌早筛检测", "送检目的", sample["purpose"]]
    ]
    info_col_widths = [70, 130, 70, 130]
    info_table = create_wrapped_table(info_data, info_col_widths, styles["TableCell"])
    info_table.setStyle(get_table_style())
    elements.append(info_table)
    elements.append(Paragraph(
        "注：本报告结果仅对本次送检样本负责，限受检者本人拆阅。",
        ParagraphStyle(
            name="NoteSmall",
            parent=styles["Body"],
            fontSize=10,
            textColor=THEME_COLOR["text"]
        )
    ))
    elements.append(Spacer(1, 20))

  
    cancer_excel_table, excel_df = read_cancer_excel_to_table_data(IMAGE_DIR, styles, gender)
    elements.append(cancer_excel_table)


    elements.append(Paragraph("二、检测说明", styles["Header2"]))
    elements.append(Paragraph(
        "本次检测覆盖20种常见癌种，通过排查128511种恶性突变，客观反映细胞异变状态。以下为各癌种检测基因数及特异位点数量统计：",
        styles["Body"]
    ))

    # 定义两张图片的路径（images文件夹下）
    cancer_gene_img_path = IMAGE_DIR / "基因.png"
    cancer_specific_site_img_path = IMAGE_DIR / "位点.png"

    # 插入第一张图：癌症基因.png（适配页面宽度90%，高度不超过页面40%）
    if os.path.exists(cancer_gene_img_path):
        cancer_gene_img = insert_image_with_fit(
            str(cancer_gene_img_path),
            max_width=CONTENT_WIDTH * 0.9,
            max_height=CONTENT_HEIGHT * 0.4
        )
        elements.append(cancer_gene_img)
        # 图片说明
        elements.append(Paragraph("图1：各癌种检测基因数量统计", ParagraphStyle(
            name="ImageCaption",
            parent=styles["Body"],
            fontSize=9,
            alignment=1,
            spaceAfter=15
        )))
    else:
        elements.append(Paragraph("【癌症基因.png 图片缺失】", ParagraphStyle(
            name="ImageMissingNote",
            fontName=FONT,
            fontSize=10,
            textColor=THEME_COLOR["risk_high"],
            alignment=1,
            spaceAfter=15
        )))

    # 插入第二张图：癌症特异位点.png（适配页面宽度90%，高度不超过页面40%）
    if os.path.exists(cancer_specific_site_img_path):
        cancer_specific_site_img = insert_image_with_fit(
            str(cancer_specific_site_img_path),
            max_width=CONTENT_WIDTH * 0.9,
            max_height=CONTENT_HEIGHT * 0.4
        )
        elements.append(cancer_specific_site_img)
        # 图片说明
        elements.append(Paragraph("图2：各癌种特异位点数量统计", ParagraphStyle(
            name="ImageCaption",
            parent=styles["Body"],
            fontSize=9,
            alignment=1,
            spaceAfter=15
        )))
    else:
        elements.append(Paragraph("【癌症特异位点.png 图片缺失】", ParagraphStyle(
            name="ImageMissingNote",
            fontName=FONT,
            fontSize=10,
            textColor=THEME_COLOR["risk_high"],
            alignment=1,
            spaceAfter=15
        )))
        
        
#    elements.append(PageBreak())
 
    # 第二章节：检测结果分析
    add_chapter_cover(elements, "第二章节：检测结果分析", "章节2封面.png", styles)
    elements.append(Paragraph("第二章节：检测结果分析", styles["Header1"]))
    elements.append(Paragraph("一、基因变异检测结果", styles["Header2"]))
    elements.append(Paragraph(
        "以下为检测到的特异性突变位点，详细风险评估见第五部分：",
        styles["Body"]
    ))
    
    variant_data = [
        ["编码", "基因", "突变位点(HGVS)", "相关癌种", 
         "COSMIC阳性", "COSMIC总样本",  "panel样本数","正常人AF"],  # 表头变短
    ]
    
    # 加载cosmic频率映射表
    cosmic_tsv_path = TABLE_DIR / "all_COSO_mutation_sample_freq_new.tsv"
    hgvs_freq_map = load_cosmic_frequency_map(sample["name"], cosmic_tsv_path)
    
    # 先定义panel样本数列表（按顺序对应4个位点）
    panel_sample_nums = [2499, 2499, 2345, 1490]
    # 遍历变异数据
    for idx, (_, r) in enumerate(variants.iterrows(), 1):
        hgvs = r["hgvs"]
        cosmic_info = hgvs_freq_map.get(hgvs, {"total_sample": 0, "positive_sample": 0, "frequency": 0.0})
        positive_sample = cosmic_info["positive_sample"]
        total_sample = cosmic_info["total_sample"]
    
        normal_info = call_bash_get_normal_af(
            hgvs_site=hgvs,
            normal_tsv_path=NORMAL_COSMIC_TSV_PATH,
            bash_script_path="./get_normal_af.sh"
        )
        normal_af = f"{normal_info['af']:.2e}"  # 科学计数法，更省空间
            # 新增：获取对应索引的panel样本数（超出列表长度时默认0）
        panel_num = panel_sample_nums[idx-1] if (idx-1) < len(panel_sample_nums) else 0
    
        variant_data.append([
            str(idx),
            r["gene"],
            r["hgvs"],
            r["cancer"],
            str(positive_sample),
            str(total_sample),
            str(panel_num),  # 新增panel样本数列值
            normal_af
        ])
    
    # ====================== 关键：压缩列宽，保证不溢出 ======================
    variant_col_widths = [30, 50, 150, 85, 50, 55, 60, 65]
    # =====================================================================
    
    variant_table = create_wrapped_table(
        variant_data,
        variant_col_widths,
        styles["TableCell"]
    )
    variant_table.setStyle(get_table_style())
    elements.append(variant_table)
    elements.append(Spacer(1, 15))   
   
    
    # 检测结果综述（核心修改：移除原Excel表格引用逻辑，仅保留突变比例说明）
    elements.append(Paragraph("二、检测结果综述", styles["Header2"]))

    # 突变比例说明（核心修复：使用reportlab支持的<sup>标签实现上标，正确显示10的次方）
    elements.append(Paragraph("突变比例与细胞量关系", styles["Header2"]))
    # 修复点1：使用<sup>标签包裹次方数，实现上标显示（reportlab的Paragraph支持该HTML标签）
    ratio_data = [
        ["突变比例", "相当于体内变异细胞数量", "临床意义"],
        ["1%", f"~10<sup>9</sup> 个", "高负荷突变，需紧急就医排查"],
        ["0.1%", f"~10<sup>8</sup> 个", "中负荷突变，建议1个月内复查"],
        ["0.01%", f"~10<sup>7</sup> 个", "低负荷突变，建议3-6个月复查"],
        ["0.05%-0.08%", f"~10<sup>7</sup>-10<sup>8</sup> 个", "本次检测部分位点突变范围"]
    ]
    ratio_col_widths = [80, 150, 170]
    ratio_table = create_wrapped_table(ratio_data, ratio_col_widths, styles["TableCell"])
    ratio_table.setStyle(get_table_style())
    elements.append(ratio_table)
    elements.append(PageBreak())
    # 插入图片（自动适配页面宽度的90%，高度不超过页面40%）

    elements.append(Paragraph("人群突变位点中位数图", ParagraphStyle(
        name="ImageCaption",
        parent=styles["Body"],
        fontSize=9,
        alignment=1,
        spaceAfter=15
    )))

    tbm_img = insert_image_with_fit(
        str(TMB_IMAGE_PATH),
        max_width=CONTENT_WIDTH * 0.9,
        max_height=CONTENT_HEIGHT * 0.4
    )
    elements.append(tbm_img)

    # 新增：插入TMB2.png图片
    tbm2_img = insert_image_with_fit(
        str(TMB2_IMAGE_PATH),
        max_width=CONTENT_WIDTH * 0.9,
        max_height=CONTENT_HEIGHT * 0.4
    )
    elements.append(tbm2_img)

    elements.append(PageBreak())
    # =========================
    # 核心改动：第五部分 - 特异突变点排查结果（读取Excel数据，映射对应列）
    # =========================
    elements.append(Paragraph("三、特异突变点排查结果", styles["Header2"]))
    elements.append(Paragraph(
        "以下为各组织特异突变点排查结果，红绿风险条直观展示癌化风险（红=高风险，绿=低风险）：",
        styles["Body"]
    ))

    # ========== 核心：从TSV数据中统计各癌种位点数量（无映射，直接按原始名称匹配） ==========
    cancer_site_stat = count_cancer_specific_sites(variants)

    # ========== 核心修改：从Excel数据构建表格（替代硬编码的base_cancer_data） ==========
    screen_raw_data = [["编码", "患癌风险组织", "检测评估结果", "检测基因数", "检测位点数", "癌化风险"]]

    if excel_df is not None and not excel_df.empty:
        # 查找Excel中的关键列（模糊匹配）
        cancer_col = None  # 癌种名称列
        gene_count_col = None  # 检测基因数量列
        site_count_col = None  # 检测位点数量列

        # 遍历Excel列名，匹配关键列
        # 修改后（适配新Excel列名）
        for col in excel_df.columns:
            col_name = str(col).lower()
            # 匹配癌种列（不变）
            if any(key in col_name for key in ["癌种", "癌症", "肿瘤", "组织"]):
                cancer_col = col
            # 匹配「特异位点-基因数」列
            elif any(key in col_name for key in ["特异位点", "gene count"]):
                gene_count_col = col
            # 匹配「特异位点-位点数」列
            elif any(key in col_name for key in ["位点数", "site count"]):
                site_count_col = col

        # 兜底：如果未找到匹配列，使用前3列作为默认
        if cancer_col is None:
            cancer_col = excel_df.columns[0] if len(excel_df.columns) > 0 else ""
        if gene_count_col is None:
            gene_count_col = excel_df.columns[1] if len(excel_df.columns) > 1 else ""
        if site_count_col is None:
            site_count_col = excel_df.columns[2] if len(excel_df.columns) > 2 else ""

        # 遍历Excel数据行构建表格
        for idx, (_, row) in enumerate(excel_df.iterrows(), 1):
            cancer_name = str(row[cancer_col]) if cancer_col else ""
            gene_count = str(row[gene_count_col]) if gene_count_col else ""
            total_sites = str(row[site_count_col]) if site_count_col else ""

            # 匹配TSV中的位点统计
            site_count = cancer_site_stat.get(cancer_name, 0)
            if site_count > 0:
                detection_result = f"特异性突变点{site_count}个"
                # 根据突变数量判断风险等级
                if site_count >= 10:
                    risk_level = "阳（高风险）"
                elif site_count >= 3:
                    risk_level = "阳（中风险）"
                else:
                    risk_level = "阳（低风险）"
            else:
                detection_result = "未检测到明确突变位点"
                risk_level = "阴（低风险）"

            # 添加到表格数据
            screen_raw_data.append([
                str(idx),  # 编码
                cancer_name,  # 患癌风险组织
                detection_result,  # 检测评估结果
                gene_count,  # 检测基因数（映射Excel的检测基因数量）
                total_sites,  # 检测位点数（映射Excel的检测位点数量）
                risk_level  # 癌化风险
            ])
    else:
        # Excel读取失败时使用默认数据
        default_data = [
            ["1", "肺腺癌", "特异性突变点15个", "1132", "333030", "阳（高风险）"],
            ["2", "肺癌", "特异性突变点12个", "1122", "249188", "阳（高风险）"],
            ["3", "乳腺癌", "未检测到明确突变位点", "1109", "156910", "阴（低风险）"],
            ["4", "肝癌", "特异性突变点8个", "1087", "111080", "阳（中风险）"],
            ["5", "大肠癌", "特异性突变点5个", "1135", "513862", "阳（低风险）"]
        ]
        # 按性别过滤默认数据
        default_data = filter_cancer_by_gender(default_data, sample["gender"])
        screen_raw_data.extend(default_data)

    # ========== 风险值映射（基于Excel数据计算） ==========
    risk_calculation_map = {}
    for row_idx in range(1, len(screen_raw_data)):
        raw_row = screen_raw_data[row_idx]
        code = int(raw_row[0])
        # 提取特异位点数量（从检测评估结果中解析）
        detection_result = raw_row[2]
        if "特异性突变点" in detection_result and "个" in detection_result:
            specific_sites = int(detection_result.replace("特异性突变点", "").replace("个", ""))
        else:
            specific_sites = 0

        # 提取总位点数量
        try:
            total_sites = int(raw_row[4])
        except (ValueError, IndexError):
            total_sites = 0

        # 计算突变比例和风险值
        if total_sites == 0:
            mutation_ratio = 0
            risk_value = 0.0
        else:
            mutation_ratio = specific_sites / total_sites
            risk_value = calculate_cancer_risk(specific_sites, total_sites)

        # 判定风险文本
        if mutation_ratio >= 0.00004:
            risk_text = "高风险"
        elif mutation_ratio >= 0.00001:
            risk_text = "中风险"
        else:
            risk_text = "低风险"

        risk_calculation_map[code] = {
            "specific": specific_sites,
            "total": total_sites,
            "text": risk_text,
            "value": risk_value
        }

    # ========== 构建带风险条的表格数据 ==========
    screen_data = [screen_raw_data[0]]  # 表头
    temp_image_paths = []  # 存储临时图片路径

    for row_idx in range(1, len(screen_raw_data)):
        raw_row = screen_raw_data[row_idx]
        code = int(raw_row[0])  # 编码
        risk_config = risk_calculation_map[code]

        # 获取风险值和风险文本
        risk_value = risk_config["value"]
        risk_text = risk_config["text"]

        # 生成风险条图片
        # 生成风险条图片（传入表格列宽，精准匹配）
        risk_img_path = generate_risk_bar_image(
            risk_value=risk_value,
            risk_text=risk_text,
            temp_dir=TEMP_DIR,
            table_col_width=110  # 与screen_col_widths中风险列宽度一致
        )

        # 直接加载PDF矢量图，尺寸与画布完全一致，无溢出
        risk_img = Image(risk_img_path)
        # 尺寸与表格列宽、画布高度完全匹配，不超界
        risk_img.drawWidth = 110
        risk_img.drawHeight = 30
        risk_img.hAlign = "CENTER"

        # 替换风险列为图片对象
        new_row = raw_row[:5] + [risk_img]
        screen_data.append(new_row)

    # ========== 生成表格并添加到PDF ==========
    # 表格列宽度（第六列加宽以适配风险条）
    screen_col_widths = [40, 120, 130, 70, 90, 110]

    # 创建表格（指定第5列（索引5）为风险图片列）
    screen_table = create_wrapped_table(
        screen_data,
        screen_col_widths,
        styles["TableCell"],
        risk_image_cols=[5]  # 第六列（索引5）是风险图片
    )
    screen_table.setStyle(get_table_style())

    elements.append(screen_table)

    # =========================
    # 新增功能：突变癌种筛查结果分述
    # =========================
    # cosmic_tsv_path = TABLE_DIR / "all_COSO_mutation_sample_freq.tsv"
    # generate_beautiful_pdf函数中
# 提前加载cosmic频率映射表
    cosmic_tsv_path = TABLE_DIR / "all_COSO_mutation_sample_freq_new.tsv"
    hgvs_freq_map = load_cosmic_frequency_map(sample["name"], cosmic_tsv_path)
    
    # 生成突变癌种筛查结果分述时传递hgvs_freq_map
    generate_cancer_screening_narration(elements, styles, cancer_site_stat, excel_df, gender, sample["name"],
                                        cosmic_tsv_path, hgvs_freq_map)

    # 干预建议（适配肺癌）
#    elements.append(PageBreak())
    # 第三章节：医学解读与建议
    add_chapter_cover(elements, "第三章节：医学解读与建议", "章节3封面.png", styles)
    elements.append(Paragraph("第三章节：医学解读与建议", styles["Header1"]))
    elements.append(Paragraph("一、健康干预建议", styles["Header2"]))
    # 通用高风险建议
    elements.append(Paragraph("1. 高风险癌种通用干预建议", styles["Header2"]))
    advice_data1 = [
        ["干预类型", "具体建议"],
        ["就医建议", "3个月内完成对应癌种的专项影像学/病理学检查"],
        ["生活防范", "避免烟酒、熬夜、高脂饮食等高危因素，保持规律作息"],
        ["膳食保护", "增加新鲜蔬果摄入，减少加工肉类、腌制食品食用"],
        ["定期复查", "首次复查后，每3-6个月进行一次ctDNA+肿瘤标志物检测"]
    ]
    advice_col_widths = [80, 320]
    advice_table1 = create_wrapped_table(advice_data1, advice_col_widths, styles["TableCell"])
    advice_table1.setStyle(get_table_style())
    elements.append(advice_table1)

    # 中风险建议
    elements.append(Paragraph("2. 中风险癌种干预建议", styles["Header2"]))
    advice_data2 = [
        ["干预类型", "具体建议"],
        ["就医建议", "6个月内完成对应癌种的常规筛查"],
        ["生活防范", "控制体重（BMI<25），适度运动（每周≥150分钟中等强度运动）"],
        ["膳食保护", "针对性摄入防癌食物（如肺腺癌：百合、银耳；肝癌：枸杞、蒲公英）"],
        ["定期复查", "每年一次专项筛查，每6个月监测血糖/血压/血脂"]
    ]
    advice_table2 = create_wrapped_table(advice_data2, advice_col_widths, styles["TableCell"])
    advice_table2.setStyle(get_table_style())
    elements.append(advice_table2)
    elements.append(PageBreak())

    # 新增：调用癌症形成机制模块
    add_cancer_formation_section(elements, styles)

    # 新增：调用ctDNA相关模块
    add_ctdna_sections(elements, styles)

    elements.append(PageBreak())

    # 报告备注与落款（序号调整为第九部分，与新增模块衔接）
    elements.append(Paragraph("四、报告备注", styles["Header2"]))

    note_data = [
        ["备注项目", "说明"],
        ["检测技术", "基于高通量测序（NGS），检测灵敏度0.05%"],
        ["数据隐私", "您的基因组数据已加密存储，仅本人可查阅"],
        ["复查建议", "阳性突变者建议每3-6个月复查一次，阴性者每年一次"],
        ["临床参考", "本报告需结合临床症状、影像学检查综合判断"],
        ["风险算法", "癌化风险采用调整型Logistic函数计算，取值范围0-100%"]
    ]
    note_col_widths = [100, 300]
    note_table = create_wrapped_table(note_data, note_col_widths, styles["TableCell"])
    note_table.setStyle(get_table_style(template_type="message"))
    elements.append(note_table)

    # 落款
    elements.append(Spacer(1, 30))
    elements.append(Paragraph("版权所有：OncoSeek生命科技有限公司", ParagraphStyle(
        name="CompanyName",
        parent=styles["Body"],
        fontSize=12,
        alignment=1,
        bold=True,
        textColor=THEME_COLOR["primary"]
    )))

    elements.append(Paragraph(f"报告生成时间：{pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}", ParagraphStyle(
        name="GenerateTime",
        parent=styles["Body"],
        fontSize=9,
        alignment=1,
        textColor=THEME_COLOR["text"]
    )))

    # 生成PDF
    doc.build(elements)

    # 清理临时图片文件
    for img_path in temp_image_paths:
        try:
            os.remove(img_path)
        except Exception as e:
            print(f"⚠️  清理临时图片失败：{img_path} - {e}")

    print(f"✅ 美化版PDF生成成功：{output_pdf}")

# =========================
# 新增模块1：癌症形成机制（含图片）
# =========================
def add_cancer_formation_section(elements, styles):
    
    elements.append(Paragraph("二、癌症形成机制", styles["Header2"]))
    elements.append(Paragraph(
        "了解癌症的成因，避免误解造成恐慌：",
        styles["Body"]
    ))
    elements.append(Paragraph(
        "2016年6月，世界卫生组织等国际机构把癌症重新定义为可以治疗、控制甚至治愈的慢性病。癌症疾病是慢性的过程，从正常细胞演变成癌细胞，再形成肿瘤，通常需要10-20年甚至更长。当危险因素对机体防御体系损害严重，机体修复能力下降，细胞内基因突变累积到一定程度，癌症才会发生，癌症的发生大都是按照癌前病变、原位癌、浸润癌、转移癌的过程发展。癌前病变阶段体内含有突变的细胞数量一般在10<sup>8</sup>以下。这个阶段属于累计期，低风险，但影像学还无法检测，需要借助DNA检测的新技术手段。",
        styles["Body"]
    ))
    elements.append(Paragraph(
        "癌症成因：化学刺激、辐射刺激、病毒细菌等因素导致DNA损伤，正常细胞受损后，经10-20年累积突变，先进入癌症超早期（DNA检测可检出，肿瘤标志物和影像学不能检出），再到癌症早期（肿瘤标志物、影像学等手段极难检出），最后到癌症中晚期（PET-CT等影像学手段可检出），之后3个月-3年可能快速增长、转移，导致死亡。",
        styles["Body"]
    ))

    # 插入癌症成因图片（自动适配页面）
    if os.path.exists(CANCER_FORMATION_IMAGE_PATH):
        cancer_img = insert_image_with_fit(str(CANCER_FORMATION_IMAGE_PATH), max_height=CONTENT_HEIGHT * 0.5)
        elements.append(cancer_img)
        elements.append(Paragraph("图1：癌症形成机制示意图", ParagraphStyle(
            name="ImageCaption",
            parent=styles["Body"],
            fontSize=9,
            alignment=1,
            spaceAfter=15
        )))
    else:
        elements.append(Paragraph("【癌症成因示意图缺失】", ParagraphStyle(
            name="ImageMissing",
            parent=styles["Body"],
            fontSize=10,
            textColor=THEME_COLOR["risk_high"],
            alignment=1
        )))
    elements.append(Spacer(1, 15))


# =========================
# 新增模块2：ctDNA相关内容（解读+共识+特点+临床意义+适用人群）
# =========================
def add_ctdna_sections(elements, styles):
    # 1. ctDNA解读
    elements.append(PageBreak())

    elements.append(Paragraph("三、cfDNA解读", styles["Header2"]))
    elements.append(Paragraph(
        "循环肿瘤DNA（circulating tumor DNA，ctDNA）：循环肿瘤DNA（circulating tumor DNA，ctDNA），是一种存在于血浆或血清、脑脊液等体液中的细胞外DNA，主要来自于坏死或凋亡的肿瘤细胞、肿瘤细胞分泌的外排体及循环肿瘤细胞，大小通常为160-180bp。ctDNA是游离DNA（cell-free DNA，cfDNA）中的一类，所占比例低（0.1%-1%之间），因此检测难度较大。二代测序（NGS）技术的成熟，提高了ctDNA检测的灵敏度和准确度，加速推进ctDNA检测应用于临床，协助医生进行个体化诊疗，为肿瘤患者的治疗带来极大的便利。",
        styles["Body"]
    ))
    elements.append(Paragraph(
        "循环肿瘤DNA（ctDNA）无创个体化诊疗基因检测，通过抽取外周血获取ctDNA进行测序，分析肿瘤药物相关基因变异情况，全面、准确解读肿瘤药物与基因的关联，为医生制定个体化用药诊疗方案提供帮助。",
        styles["Body"]
    ))
    elements.append(Spacer(1, 10))

    # # 2. ctDNA共识（表格形式，与PDF一致）
    elements.append(Paragraph("（一）、ctDNA共识", styles["Header3"]))


    elements.append(Paragraph(
        "1. 高通量测序敏感性、特异性、准确性高，且组织样本与ctDNA液体活检一致性较高，ctDNA检测可作为组织活检的补充；<br/>"
            "2. ctDNA检测可无创、实时取样，可以作为复发及耐药动态检测的有利工具；<br/>"
            "3. ctDNA检测可克服肿瘤组织异质性带来的检测差异；<br/>"
            "4. 高通量测序基因检测可发现新的驱动基因，助推新药研发。<br/>",
        styles["Body"]
    ))
    # 3. 循环肿瘤DNA检测特点（表格形式）
    # 循环肿瘤DNA检测特点
    elements.append(Paragraph("（二）循环肿瘤DNA检测特点", styles["Header3"]))

    elements.append(Paragraph(

          "1. 取样方便<br/>ctDNA检测只需抽取血液即可完成对肿瘤细胞DNA的分析与解读，避免穿刺、手术等有创方式采集肿瘤组织的风险和痛苦。在恶性肿瘤发生早期，ctDNA含量和基因的异常改变就可以被检测到。ctDNA检测可以实现实时、多阶段、个体化诊疗服务。<br/>"
            "2. 检测技术成熟<br/>利用NGS检测ctDNA，灵敏准确，可检测到低至0.1%的低频突变。Dawson等人在《新英格兰医学》发表的一项研究中指出：与其他类型的肿瘤标记物比较，ctDNA具有更高的敏感性，可以更早的反应肿瘤的进展情况和治疗效果。ctDNA携带有肿瘤患者的基因信息，定量或定性分析这些循环DNA对肿瘤的早期诊断、治疗、病情监测及预后的评价具有重要的临床价值。ctDNA检测能够为肿瘤个体化治疗呈现较全面的分子生物学特性，是一种理想的肿瘤标志物。<br/>"
            "3. 应用范围广<br/>几乎所有肿瘤细胞都可以将DNA片段释放到血液中，因此，ctDNA能够体现患者体内肿瘤的综合情况。临床上，ctDNA检测可应用于A）肿瘤的早期诊断；B）监控肿瘤的演化和适应性改变；C）实时监控用药治疗效果，追踪肿瘤转归、转移和复发等；D）个性化的治疗方案指导。异质性对肿瘤组织基因检测影响较大，ctDNA检测可以避免因肿瘤组织异质性导致取样的随机性和局部性，使得肿瘤突变结果更全面、可靠。<br/>"
            ,
        styles["Body"]
    ))
    # 临床意义
    elements.append(Paragraph("（三）临床意义", styles["Header3"]))

    elements.append(Paragraph(


            "1. 个体化用药参考：检测药物靶点基因，解读相关位点，分析基因变异与药物的关系，根据肿瘤患者特有的基因变异，评估药物的疗效和相关毒副作用，协助医生制定个性化的治疗方案；<br/>"
            "2. 耐药性分析：耐药性的产生源于肿瘤相关基因的不断突变，ctDNA检测的优势在于实时动态检测药物点突变，分析耐药性产生的根源，协助医生适时调整治疗方案、选择合适的药物；<br/>"
            "3. 动态检测：ctDNA检测可以多个时间点取样，协助医生实现对患者用药疗效、复发及转移的动态监测。<br/>"
        ,
        styles["Body"]
    ))
    # 5. 适用人群（表格形式）
    elements.append(Paragraph("（四）适用人群", styles["Header3"]))


    elements.append(Paragraph(

    "1. 正常体检人群 <br/>"
        "2. 40岁以上人群<br/>"  
        "3. 癌症家族史人群<br/>"
        "4. 重度饮酒，吸烟，熬夜人群<br/>"
        "5. 暴饮暴食，三餐不定人群<br/>"
        "6. 消化道，呼吸道慢性疾病人群<br/>"
        "7. 胃病，结直肠癌等癌症的高危人群<br/>"
        "8. 患有甲状腺结节，乳腺结节，肺结节，胃炎，胃肠息肉的人群<br/>"
    ,
    styles["Body"]
))
# =========================
# 主入口（支持命令行传参）
# =========================
def main(patient_name):  # 移除cancer_type参数
    """
    主函数：生成PDF报告（癌种从映射文件自动获取）
    :param patient_name: 患者姓名
    """
    # 读取患者信息
    sample = read_patient_info_tsv(patient_name)
    
    # 读取位点数据（自动关联癌种）
    variants = read_variant_tsv(patient_name)
    
    # 生成PDF
    out_pdf = OUTPUT_DIR / f"{patient_name}_美化版防癌早筛报告.pdf"
    generate_beautiful_pdf(sample, variants, out_pdf)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="生成美化版防癌早筛PDF报告（新HGVS+癌种映射格式）")
    parser.add_argument("--patient", "-p", type=str, default="张小", help="患者姓名（如：张三）")
    # 移除--cancer参数
    args = parser.parse_args()
    main(args.patient)
