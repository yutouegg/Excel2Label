import streamlit as st
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO


def create_label(data):
    """创建单个标签的数据"""
    return [
        ['供应商：横店华达彩印'],
        [f'合同号：{data["合同序号*"]}', f'计划号：{data["计划号"]}'],
        [f'送达仓库：{data["送达仓库"]}'],
        [f'材料货号：{data["材料货号*"]}'],
        [f'物料名称：{data["物料名称"]}'],
        [f'采购数量：{data["采购数量*"]}', f'交货日期：{data["交货日期"]}'],
        ['RoHS']
    ]


def create_pdf(df, labels_per_row=2):
    """生成包含所有标签的PDF文件"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=30
    )

    # 创建所有标签
    elements = []
    all_labels = []
    current_row = []

    # 计算单个标签的宽度
    label_width = (A4[0] - 60) / labels_per_row  # 60是左右边距的总和

    # 为每行数据创建标签
    for idx, row in df.iterrows():
        label_data = create_label(row)
        label_table = Table(label_data, colWidths=[label_width / 2] * 2)
        label_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))

        current_row.append(label_table)

        # 当达到每行标签数量限制时，添加到主表格
        if len(current_row) == labels_per_row:
            all_labels.append(current_row)
            current_row = []

    # 处理最后一行不完整的情况
    if current_row:
        while len(current_row) < labels_per_row:
            current_row.append('')  # 添加空白占位
        all_labels.append(current_row)

    # 创建主表格
    main_table = Table(all_labels, colWidths=[label_width] * labels_per_row)
    main_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
    ]))

    elements.append(main_table)
    doc.build(elements)

    buffer.seek(0)
    return buffer


def main():
    st.title("材料标签生成器")

    # 上传文件
    uploaded_file = st.file_uploader("上传Excel文件", type=['xlsx', 'xls'])

    # 设置每行标签数量
    labels_per_row = st.number_input("每行标签数量", min_value=1, max_value=3, value=2)

    if uploaded_file is not None:
        # 读取Excel文件
        df = pd.read_excel(uploaded_file)

        # 过滤出需要的列
        needed_cols = ['合同序号*', '计划号', '送达仓库', '材料货号*', '物料名称', '采购数量*', '交货日期']

        # 过滤非空行
        df_filtered = df[df['合同序号*'].notna()]

        # 显示数据预览
        st.write(f"共找到 {len(df_filtered)} 条记录")
        st.write("数据预览：")
        st.dataframe(df_filtered[needed_cols].head())

        # 生成PDF按钮
        if st.button("生成PDF"):
            pdf = create_pdf(df_filtered, labels_per_row)
            st.download_button(
                label="下载PDF文件",
                data=pdf,
                file_name="material_labels.pdf",
                mime="application/pdf"
            )


if __name__ == "__main__":
    main()