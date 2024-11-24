import pandas as pd
import streamlit as st
import base64
from io import BytesIO

# Streamlit 设置
st.title("Excel 标签生成工具")
st.write("上传包含采购信息的 Excel 表")

# 上传 Excel 文件
uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])

# 定义所需列（修改为实际的列名）
needed_cols = ['完成', '合同序号*', '计划号', '送达仓库', '材料货号*', '物料名称', '采购数量*', '交货日期']

if uploaded_file:
    try:
        # 读取 Excel 数据
        df = pd.read_excel(uploaded_file)
        # 格式化交货日期为仅包含年月日
        if '交货日期' in df.columns:
            df['交货日期'] = pd.to_datetime(df['交货日期']).dt.strftime('%Y-%m-%d')

        # 过滤需要的列并删除全空行
        if all(col in df.columns for col in needed_cols):
            filtered_df = df[needed_cols].dropna(how='all')

            st.write("处理后的小表格数据：")
            st.dataframe(filtered_df)


            # 生成 HTML 表格，适配 A4 纸张
            def generate_html_labels(df):
                html_content = """
                <html>
                <head>
                <style>
                    @page {
                        size: A4;
                        margin: 10mm;
                    }
                    body {
                        font-family: Arial, sans-serif;
                        margin: 0;
                        padding: 0;
                    }
                    .page {
                        width: 210mm;
                        height: 297mm;
                        padding: 10mm;
                        box-sizing: border-box;
                        page-break-after: always;
                    }
                    .label-container {
                        display: grid;
                        grid-template-columns: repeat(2, 1fr);
                        grid-template-rows: repeat(7, 1fr);
                        gap: 5mm;
                        height: 100%;
                    }
                    .label {
                        border: 1px solid black;
                        padding: 2mm;
                        width: 90mm;
                        height: 35mm;
                        box-sizing: border-box;
                        page-break-inside: avoid;
                        font-size: 8pt;
                        display: flex;
                        flex-direction: column;
                        position: relative;
                    }
                    .info-row {
                        display: flex;
                        justify-content: space-between;
                        align-items: flex-start;
                        margin-bottom: 1mm;
                        flex-wrap: wrap;
                    }
                    .info-item {
                        flex: 1 1 auto;
                        min-width: 45%;
                        margin-bottom: 1mm;
                        line-height: 1.2;
                    }
                    .material-name {
                        width: 100%;
                        word-break: break-all;
                        white-space: normal;
                        line-height: 1.2;
                        margin-bottom: 1mm;
                        margin-top: 1mm;
                    }
                    .contract-no {
                        font-weight: bold;
                        margin-bottom: 1mm;
                        font-size: 9pt;
                    }
                    .supplier {
                        position: absolute;
                        bottom: 2mm;
                        right: 2mm;
                        font-size: 7pt;
                    }
                </style>
                </head>
                <body>
                """

                # 计算需要的页数
                labels_per_page = 14  # 2 × 7
                total_pages = (len(df) + labels_per_page - 1) // labels_per_page

                for page in range(total_pages):
                    html_content += '<div class="page"><div class="label-container">'
                    start_idx = page * labels_per_page
                    end_idx = min((page + 1) * labels_per_page, len(df))

                    for idx in range(start_idx, end_idx):
                        row = df.iloc[idx]
                        html_content += '<div class="label">'

                        # 合同编号行
                        html_content += f'<div class="contract-no">合同编号：{row["完成"]}</div>'

                        # 第一行信息
                        html_content += '<div class="info-row">'
                        html_content += f'<div class="info-item">序号：{row["合同序号*"]}</div>'
                        html_content += f'<div class="info-item">计划号：{row["计划号"]}</div>'
                        html_content += f'<div class="info-item">采购数量：{row["采购数量*"]}</div>'
                        html_content += '</div>'

                        # 第二行信息
                        html_content += '<div class="info-row">'
                        html_content += f'<div class="info-item">送达仓库：{row["送达仓库"]}</div>'
                        html_content += f'<div class="info-item">材料货号：{row["材料货号*"]}</div>'
                        html_content += '</div>'

                        # 物料名称（单独一行，允许换行）
                        html_content += f'<div class="material-name">物料名称：{row["物料名称"]}</div>'

                        # 交货日期
                        html_content += f'<div class="info-row">'
                        html_content += f'<div class="info-item">交货日期：{row["交货日期"]}</div>'
                        html_content += '</div>'

                        # 供应商信息
                        html_content += '<div class="supplier">供应商：横店华达彩印</div>'
                        html_content += "</div>"  # 结束 label

                    # 填充剩余空白标签以保持布局
                    for _ in range(end_idx - start_idx, labels_per_page):
                        html_content += '<div class="label"></div>'

                    html_content += "</div></div>"

                html_content += "</body></html>"
                return html_content


            # 调用函数生成 HTML
            html_labels = generate_html_labels(filtered_df)


            # 提供 HTML 文件下载
            def download_html(data, filename):
                b64 = base64.b64encode(data.encode()).decode()
                href = f'<a href="data:text/html;base64,{b64}" download="{filename}">点击下载标签文件</a>'
                return href


            st.markdown(download_html(html_labels, "labels.html"), unsafe_allow_html=True)

        else:
            st.error("上传的 Excel 文件中缺少所需的列，请检查！")
    except Exception as e:
        st.error(f"处理文件时出错：{e}")