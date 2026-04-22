import io
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
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
        raise ValueError(
            f"读取到的列数为 {df.shape[1]}，但预期应为 10 列（G:P）。"
        )

    headers = list(df.columns)
    name_col = headers[0]
    metric_cols = headers[1:]

    if len(metric_cols) != METRIC_COUNT:
        raise ValueError(
            f"数值指标列数量为 {len(metric_cols)}，但预期应为 {METRIC_COUNT}。"
        )

    # 清洗姓名列
    df[name_col] = df[name_col].fillna("").astype(str).str.strip()

    # 数值列转为数值，非数值按 0 处理
    for col in metric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # 只保留固定名单中的人员
    df_filtered = df[df[name_col].isin(FIXED_LIST)].copy()

    # 分组求和
    if not df_filtered.empty:
        grouped = (
            df_filtered.groupby(name_col, as_index=False)[metric_cols]
            .sum()
        )
    else:
        grouped = pd.DataFrame(columns=[name_col] + metric_cols)

    # 按固定名单顺序输出，缺失人员补 0
    result_df = pd.DataFrame({name_col: FIXED_LIST})
    result_df = result_df.merge(grouped, on=name_col, how="left")
    for col in metric_cols:
        result_df[col] = result_df[col].fillna(0)

    # 尽量把整数型结果显示为整数
    for col in metric_cols:
        series = result_df[col]
        if (series % 1 == 0).all():
            result_df[col] = series.astype("int64")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="处理结果")
        workbook = writer.book
        worksheet = writer.sheets["处理结果"]

        # 样式
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

        # 标题行
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_alignment
            cell.border = thin_border

        # 数据区
        for row in worksheet.iter_rows(min_row=2, max_row=max_row, min_col=1, max_col=max_col):
            for cell in row:
                cell.alignment = center_alignment
                cell.border = thin_border

        # 隔行底色
        for row_idx in range(2, max_row + 1, 2):
            for col_idx in range(1, max_col + 1):
                worksheet.cell(row=row_idx, column=col_idx).fill = zebra_fill

        # 自动列宽
        for col in worksheet.columns:
            max_length = 0
            column_letter = col[0].column_letter
            for cell in col:
                value = "" if cell.value is None else str(cell.value)
                if len(value) > max_length:
                    max_length = len(value)
            worksheet.column_dimensions[column_letter].width = max_length + 2

    output.seek(0)
    return output.getvalue(), result_df


def build_output_filename(uploaded_name: str) -> str:
    path = Path(uploaded_name)
    stem = path.stem or "output"
    return f"{stem}_处理结果.xlsx"


st.title("Excel 汇总工具")
st.caption("上传 Excel 文件后，自动读取 Sheet1 的 G:P 区域，按固定名单汇总，并生成新的结果 Excel。")

with st.expander("固定名单", expanded=False):
    st.write("、".join(FIXED_LIST))

uploaded_file = st.file_uploader(
    "上传 Excel 文件",
    type=["xlsx", "xlsm"],
    accept_multiple_files=False,
    help="要求工作表名为 Sheet1，G列为姓名，H:P 为 9 个数值指标。",
)

sheet_name = st.text_input("工作表名称", value=DEFAULT_SHEET_NAME)

if uploaded_file is not None:
    st.info(f"已上传文件：{uploaded_file.name}")

    if st.button("开始处理", type="primary"):
        try:
            excel_bytes, result_df = process_excel(uploaded_file, sheet_name=sheet_name)
            output_filename = build_output_filename(uploaded_file.name)

            st.success("处理完成。你可以先预览结果，再下载 Excel 文件。")
            st.dataframe(result_df, use_container_width=True)

            st.download_button(
                label="下载结果 Excel",
                data=excel_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as exc:
            st.error(str(exc))
else:
    st.write("请先上传一个 Excel 文件。")
