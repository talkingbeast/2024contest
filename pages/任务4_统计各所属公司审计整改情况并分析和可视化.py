import pandas as pd
import streamlit as st
import tkinter as tk
from tkinter import filedialog
import plotly.express as px
import os
from pathlib import Path

# Set up tkinter
root = tk.Tk()
root.withdraw()
root.wm_attributes('-topmost', 1)

# Streamlit UI
st.title("任务四：审计整改数据分析")
st.write('请按照要求选择汇总台账文件，然后点击“开始分析”按钮进行数据分析。')

# Initialize session state
if 'file_path' not in st.session_state:
    st.session_state.file_path = ""
if 'save_folder' not in st.session_state:
    st.session_state.save_folder = ""

# Function to select folder
def select_folder():
    folder = filedialog.askdirectory(master=root)
    if folder:
        st.session_state.save_folder = folder

# File picker for selecting the summary ledger
if st.button('选择汇总台账文件'):
    file_path = filedialog.askopenfilename(master=root, filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        st.session_state.file_path = file_path

# Display the selected file and folder paths
st.text_input('已选择汇总台账文件:', st.session_state.file_path)
st.text_input('选择保存文件夹:', st.session_state.save_folder)

# Button to select folder for saving results
if st.button('选择保存文件夹'):
    select_folder()

if st.session_state.file_path and st.session_state.save_folder:
    if st.button('开始分析'):
        # Load data
        df = pd.read_excel(st.session_state.file_path)

        # Ensure required columns exist
        required_columns = [
            '所属责任公司', '整改计划及措施', '是否整改完成', '提示问题金额（万元）'
        ]
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"数据中缺少以下必要列: {', '.join(missing_columns)}")
        else:
            # Group by company and calculate required metrics
            grouped = df.groupby('所属责任公司').agg(
                总数量=('整改计划及措施', 'count'),
                已整改数量=('是否整改完成', lambda x: (x == '是').sum()),
                未整改数量=('是否整改完成', lambda x: (x == '否').sum()),
                总金额=('提示问题金额（万元）', 'sum'),
                已整改金额=('提示问题金额（万元）', lambda x: df.loc[x.index][df['是否整改完成'] == '是']['提示问题金额（万元）'].sum()),
                未整改金额=('提示问题金额（万元）', lambda x: df.loc[x.index][df['是否整改完成'] == '否']['提示问题金额（万元）'].sum())
            )

            grouped.reset_index(inplace=True)
            st.write(f'点击表格右上角下载按钮可直接保存为csv')
            st.dataframe(grouped)

            # Save summary to Excel
            summary_path = Path(st.session_state.save_folder) / '审计整改汇总.xlsx'
            grouped.to_excel(summary_path, index=False)
            st.success(f"汇总表格已保存为: {summary_path}")

            # Create interactive charts
            fig1 = px.bar(grouped, x='所属责任公司', y=['已整改数量', '未整改数量'],
                          title='各公司审计整改统计',
                          labels={'value': '数量', 'variable': '统计类型'})
            st.plotly_chart(fig1)

            fig2 = px.bar(grouped, x='所属责任公司', y=['已整改金额', '未整改金额'],
                          title='各公司审计整改金额统计',
                          labels={'value': '金额（万元）', 'variable': '统计类型'})
            st.plotly_chart(fig2)

            # Save all charts as HTML files
            html_path1 = Path(st.session_state.save_folder) / '各公司审计整改统计.html'
            html_path2 = Path(st.session_state.save_folder) / '各公司审计整改金额统计.html'
            fig1.write_html(html_path1)
            fig2.write_html(html_path2)

            st.success(f"图表已保存为: {html_path1} 和 {html_path2}")

else:
    st.info("请选择汇总台账文件和保存文件夹，然后点击 '开始分析' 按钮进行分析。")
