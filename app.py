import io
import re
from datetime import datetime
from pathlib import Path

import matplotlib as mpl
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

st.set_page_config(page_title="Excel 汇总工具", page_icon="📊", layout="centered")

FIXED_LIST = [
    "罗鑫",
    "邓博文",
    "陈立",
    "李芝瑶",
    "张菁",
    "郑岑",
    "周钰萱",
    "丰亦舟",
    "吴筱烨",
    "范颖",
    "查泽民",
]

DEFAULT_SHEET_NAME = "Sheet1"
SOURCE_USECOLS = "G:P"  # G列姓名，H:P共9列指标
METRIC_COUNT = 9
DEFAULT_FILE_STEM = f"业绩表-{datetime.now().strftime('%Y%m%d')}"
DEFAULT_EXCEL_FILENAME = f"{DEFAULT_FILE_STEM}.xlsx"


@st.cache_resource
def get_chinese_font_prop() -> fm.FontProperties | None:
    """
    尽量寻找可用于中文渲染的字体。

    为了在 Streamlit Community Cloud 上更稳妥地显示中文，优先查找：
    1. 仓库中随应用一起提供的字体文件。
    2. 系统中常见的中文字体。

    若你想保证中文图片渲染稳定，建议把以下任一字体文件放到仓库同目录或 fonts 目录：
    - NotoSansCJKsc-Regular.otf
    - SourceHanSansCN-Regular.otf
    - SimHei.ttf
    """
    candidate_paths = [
        Path(__file__).with_name("NotoSansCJKsc-Regular.otf"),
        Path(__file__).with_name("SourceHanSansCN-Regular.otf"),
        Path(__file__).with_name("SimHei.ttf"),
        Path(__file__).parent / "fonts" / "NotoSansCJKsc-Regular.otf",
        Path(__file__).parent / "fonts" / "SourceHanSansCN-Regular.otf",
        Path(__file__).parent / "fonts" / "SimHei.ttf",
    ]
    for path in candidate_paths:
        if path.exists():
            return fm.FontProperties(fname=str(path))

    preferred_keywords = [
        "NotoSansCJK",
        "Noto Sans CJK",
        "SourceHanSans",
        "Source Han Sans",
        "WenQuanYi",
        "SimHei",
        "Microsoft YaHei",
        "PingFang",
        "Arial Unicode",
    ]

    system_fonts = fm.findSystemFonts(fontpaths=None, fontext="ttf") + fm.findSystemFonts(
        fontpaths=None, fontext="otf"
    )
    for font_path in system_fonts:
        lower_path = font_path.lower()
        if any(keyword.lower() in lower_path for keyword in preferred_keywords):
            return fm.FontProperties(fname=font_path)

    return None


def sanitize_filename(filename: str, default_suffix: str) -> str:
    """清理文件名，并确保包含指定后缀。"""
    name = (filename or "").strip()
    name = name.replace("\\", "_").replace("/", "_")
    name = re.sub(r'[<>:"|?*]', "_", name)
    if not name:
        name = f"业绩表-{datetime.now().strftime('%Y%m%d')}"

    path = Path(name)
    if path.suffix.lower() != default_suffix.lower():
        name = f"{path.stem or path.name}{default_suffix}"
    return name


def build_plot_filename(excel_filename: str) -> str:
    sanitized_excel = sanitize_filename(excel_filename, ".xlsx")
    stem = Path(sanitized_excel).stem
    return f"{stem}.png"


def format_number(value) -> str:
    """用于图片表格展示的数字格式。"""
    if pd.isna(value):
        return ""
    if isinstance(value, int):
        return f"{value:,}"
    if isinstance(value, float):
        if value.is_integer():
            return f"{int(value):,}"
        return f"{value:,.2f}".rstrip("0").rstrip(".")
    return str(value)


def dataframe_elementwise_map(df: pd.DataFrame, func) -> pd.DataFrame:
    """
    对 DataFrame 做逐元素映射，兼容新旧 pandas 版本。

    pandas >= 2.1 推荐使用 DataFrame.map()
    旧版本仍可回退到 DataFrame.applymap()
    """
    if hasattr(df, "map"):
        return df.map(func)
    return df.applymap(func)


def process_excel(uploaded_file, sheet_name: str = DEFAULT_SHEET_NAME) -> tuple[bytes, pd.DataFrame]:
    """
    读取上传的 Excel，按固定名单汇总 G:P 区域的数据，并返回输出文件二进制内容与结果 DataFrame。

    约定：
    1. 指定工作表存在，默认名为 Sheet1。
    2. G列为姓名，H:P 为 9 个数值指标。
    3. 第 1 行为表头，第 2 行开始为数据。
    """
    uploaded_file.seek(0)

    try:
        df = pd.read_excel(
            uploaded_file,
            sheet_name=sheet_name,
            usecols=SOURCE_USECOLS,
            engine="openpyxl",
        )
    except ValueError as exc:
        raise ValueError(
            f"读取失败。请确认文件中存在工作表“{sheet_name}”，且 G:P 区域可读取。原始错误：{exc}"
        ) from exc

    expected_col_count = 10
    if df.shape[1] != expected_col_count:
        raise ValueError(f"读取到的列数为 {df.shape[1]}，但预期应为 10 列（G:P）。")

    headers = list(df.columns)
    name_col = headers[0]
    metric_cols = headers[1:]

    if len(metric_cols) != METRIC_COUNT:
        raise ValueError(f"数值指标列数量为 {len(metric_cols)}，但预期应为 {METRIC_COUNT}。")

    df[name_col] = df[name_col].fillna("").astype(str).str.strip()

    for col in metric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    df_filtered = df[df[name_col].isin(FIXED_LIST)].copy()

    if not df_filtered.empty:
        grouped = df_filtered.groupby(name_col, as_index=False)[metric_cols].sum()
    else:
        grouped = pd.DataFrame(columns=[name_col] + metric_cols)

    result_df = pd.DataFrame({name_col: FIXED_LIST})
    result_df = result_df.merge(grouped, on=name_col, how="left")
    for col in metric_cols:
        result_df[col] = result_df[col].fillna(0)

    for col in metric_cols:
        series = result_df[col]
        if (series % 1 == 0).all():
            result_df[col] = series.astype("int64")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="处理结果")
        worksheet = writer.sheets["处理结果"]

        header_fill = PatternFill(fill_type="solid", fgColor="4F81BD")
        header_font = Font(bold=True, color="FFFFFF")
        center_alignment = Alignment(horizontal="center", vertical="center")
        thin_side = Side(style="thin", color="000000")
        thin_border = Border(
            left=thin_side,
            right=thin_side,
            top=thin_side,
            bottom=thin_side,
        )
        zebra_fill = PatternFill(fill_type="solid", fgColor="F2F2F2")

        max_row = worksheet.max_row
        max_col = worksheet.max_column

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_alignment
            cell.border = thin_border

        for row in worksheet.iter_rows(min_row=2, max_row=max_row, min_col=1, max_col=max_col):
            for cell in row:
                cell.alignment = center_alignment
                cell.border = thin_border

        for row_idx in range(2, max_row + 1, 2):
            for col_idx in range(1, max_col + 1):
                worksheet.cell(row=row_idx, column=col_idx).fill = zebra_fill

        for col in worksheet.columns:
            max_length = 0
            column_letter = col[0].column_letter
            for cell in col:
                value = "" if cell.value is None else str(cell.value)
                display_width = len(value) + 2
                if any("\u4e00" <= char <= "\u9fff" for char in value):
                    display_width += 2
                max_length = max(max_length, display_width)
            worksheet.column_dimensions[column_letter].width = max_length

    output.seek(0)
    return output.getvalue(), result_df


def render_table_plot(result_df: pd.DataFrame) -> bytes:
    """生成高清 PNG 表格图片。"""
    font_prop = get_chinese_font_prop()
    if font_prop is not None:
        mpl.rcParams["font.family"] = font_prop.get_name()
    mpl.rcParams["axes.unicode_minus"] = False

    display_df = result_df.copy()
    formatted_display_df = dataframe_elementwise_map(display_df, format_number)
    rows, cols = formatted_display_df.shape

    fig_width = max(12, cols * 1.8)
    fig_height = max(4.8, rows * 0.65 + 1.4)

    fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=240)
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    ax.axis("off")

    table = ax.table(
        cellText=formatted_display_df.values,
        colLabels=list(formatted_display_df.columns),
        loc="center",
        cellLoc="center",
        colLoc="center",
    )

    table.auto_set_font_size(False)
    table.set_fontsize(12)
    table.scale(1.08, 1.55)

    for (row, col), cell in table.get_celld().items():
        cell.set_edgecolor("#BFBFBF")
        cell.set_linewidth(0.8)
        cell.set_facecolor("white")
        txt = cell.get_text()
        txt.set_ha("center")
        txt.set_va("center")
        if font_prop is not None:
            txt.set_fontproperties(font_prop)
        if row == 0:
            cell.set_facecolor("#4F81BD")
            txt.set_color("white")
            txt.set_weight("bold")
        elif row % 2 == 0:
            cell.set_facecolor("#F8F8F8")

    plt.tight_layout(pad=0.8)

    output = io.BytesIO()
    fig.savefig(output, format="png", dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    output.seek(0)
    return output.getvalue()


st.title("Excel 汇总工具")
st.caption(
    "上传 Excel 文件后，自动读取 Sheet1 的 G:P 区域，按固定名单汇总，生成新的结果 Excel，并导出高清表格图片。"
)

with st.expander("固定名单", expanded=False):
    st.write("、".join(FIXED_LIST))

uploaded_file = st.file_uploader(
    "上传 Excel 文件",
    type=["xlsx", "xlsm"],
    accept_multiple_files=False,
    help="要求工作表名为 Sheet1，G列为姓名，H:P 为 9 个数值指标。",
)

sheet_name = st.text_input("工作表名称", value=DEFAULT_SHEET_NAME)
custom_excel_filename = st.text_input(
    "输出 Excel 文件名",
    value=DEFAULT_EXCEL_FILENAME,
    help="可自定义输出文件名。默认格式为 业绩表-YYYYMMDD.xlsx。",
)

if uploaded_file is not None:
    st.info(f"已上传文件：{uploaded_file.name}")

    if st.button("开始处理", type="primary"):
        try:
            excel_bytes, result_df = process_excel(uploaded_file, sheet_name=sheet_name)
            excel_filename = sanitize_filename(custom_excel_filename, ".xlsx")
            plot_filename = build_plot_filename(excel_filename)
            plot_bytes = render_table_plot(result_df)

            st.success("处理完成。你可以预览结果，并分别下载 Excel 与高清图片。")
            st.dataframe(result_df, use_container_width=True)
            st.image(plot_bytes, caption="结果表格高清图片预览", use_container_width=True)

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="下载结果 Excel",
                    data=excel_bytes,
                    file_name=excel_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with col2:
                st.download_button(
                    label="下载高清图片",
                    data=plot_bytes,
                    file_name=plot_filename,
                    mime="image/png",
                    use_container_width=True,
                )

            if get_chinese_font_prop() is None:
                st.warning(
                    "当前运行环境中未检测到明确的中文字体。若图片中的中文显示异常，请把 NotoSansCJKsc-Regular.otf、"
                    "SourceHanSansCN-Regular.otf 或 SimHei.ttf 放到应用目录或 fonts 目录中后重新部署。"
                )
        except Exception as exc:
            st.error(str(exc))
else:
    st.write("请先上传一个 Excel 文件。")
