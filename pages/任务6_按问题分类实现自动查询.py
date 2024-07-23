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
st.title("任务6：按问题分类自动查询分析")
st.write('请先选择汇总台账文件，然后选择分组类别和问题分类进行分析。')

# Session state for file paths and data
if 'file_path' not in st.session_state:
    st.session_state.file_path = None

if 'save_folder' not in st.session_state:
    st.session_state.save_folder = ""

if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame()

if 'group_col' not in st.session_state:
    st.session_state.group_col = "所属责任公司"

# Function to select file
def select_file():
    file_path = filedialog.askopenfilename(master=root, filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        st.session_state.file_path = file_path
        st.session_state.df = pd.read_excel(file_path)
        non_numeric_cols = st.session_state.df.select_dtypes(exclude=['number', 'datetime']).columns
        st.session_state.group_col = non_numeric_cols[0] if "所属责任公司" not in non_numeric_cols else "所属责任公司"

# Function to select folder
def select_folder():
    folder = filedialog.askdirectory(master=root)
    if folder:
        st.session_state.save_folder = folder

# File picker for selecting the summary ledger
if st.button('选择汇总台账文件'):
    select_file()

# Dropdown for group category
if not st.session_state.df.empty:
    non_numeric_cols = st.session_state.df.select_dtypes(exclude=['number', 'datetime']).columns
    st.session_state.group_col = st.selectbox('选择分组类别:', non_numeric_cols, index=non_numeric_cols.get_loc(st.session_state.group_col))

    # Dropdown for issue category
    issue_categories = st.session_state.df['问题分类'].unique()
    selected_issue = st.selectbox('选择问题分类:', issue_categories)

    # Filtering data based on selected issue category
    filtered_df = st.session_state.df[st.session_state.df['问题分类'] == selected_issue]

    if not filtered_df.empty:
        # Group by selected group column
        grouped = filtered_df.groupby(st.session_state.group_col).agg(
            已整改数量=('是否整改完成', lambda x: (x == '是').sum()),
            未整改数量=('是否整改完成', lambda x: (x == '否').sum()),
            已整改金额=('提示问题金额（万元）', lambda x: filtered_df.loc[x.index, '提示问题金额（万元）'].where(x == '是').sum()),
            未整改金额=('提示问题金额（万元）', lambda x: filtered_df.loc[x.index, '提示问题金额（万元）'].where(x == '否').sum())
        ).reset_index()

        # Create interactive charts
        fig1 = px.bar(grouped, x=st.session_state.group_col, y=['已整改数量', '未整改数量'], title='各公司审计整改统计', labels={'value': '数量', 'variable': '统计类型'})
        fig2 = px.bar(grouped, x=st.session_state.group_col, y=['已整改金额', '未整改金额'], title='各公司审计整改金额统计', labels={'value': '金额（万元）', 'variable': '统计类型'})

        st.plotly_chart(fig1)
        st.plotly_chart(fig2)
        st.write(f'点击表格右上角下载按钮可直接保存为csv')
        st.dataframe(grouped)

        if st.button('保存分析结果'):
            select_folder()
            if st.session_state.save_folder:
                summary_path = Path(st.session_state.save_folder) / '审计整改汇总.xlsx'
                grouped.to_excel(summary_path, index=False)
                fig1.write_html(Path(st.session_state.save_folder) / '审计整改数量统计.html')
                fig2.write_html(Path(st.session_state.save_folder) / '审计整改金额统计.html')
                st.success(f"汇总表格和图表已保存至: {st.session_state.save_folder}")

else:
    st.info("请选择汇总台账文件。")
