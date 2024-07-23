import pandas as pd
from datetime import datetime
import streamlit as st
import tkinter as tk
from tkinter import filedialog
from pathlib import Path

# Set up tkinter
root = tk.Tk()
root.withdraw()
root.wm_attributes('-topmost', 1)

# Initialize session state variables
if 'file_path' not in st.session_state:
    st.session_state.file_path = ""
if 'save_folder' not in st.session_state:
    st.session_state.save_folder = ""
if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame()

# Function to select file
def select_file():
    file_path = filedialog.askopenfilename(master=root, filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        st.session_state.file_path = file_path
        st.session_state.df = pd.read_excel(file_path)

# Function to select folder
def select_folder():
    folder = filedialog.askdirectory(master=root)
    if folder:
        st.session_state.save_folder = folder

# Streamlit UI
st.title("任务8和任务9：整改事项提取与目标提示")
st.write('请选择汇总台账文件，然后选择功能按钮进行操作。')

# File picker for selecting the summary ledger
if st.button('选择汇总台账文件'):
    select_file()

if st.session_state.file_path:
    st.text_input('已选择汇总台账文件:', st.session_state.file_path)

# Display the selected file
if not st.session_state.df.empty:
    st.write(f"已加载数据文件: {st.session_state.file_path}")

# Extract current month and upcoming month data
if st.session_state.df.empty:
    st.info("请先选择汇总台账文件。")
else:
    if st.button("提取本月完成整改事项"):
        # Extract current month data
        current_month = datetime.now().strftime('%Y-%m')
        df = st.session_state.df
        df['整改完成时间'] = pd.to_datetime(df['整改完成时间'], errors='coerce')
        current_month_df = df[df['整改完成时间'].dt.strftime('%Y-%m') == current_month]

        if not current_month_df.empty:
            st.write("**本月完成整改事项(点击表格右上角下载按钮可直接保存为csv)**")
            st.dataframe(current_month_df)

            

    if st.button("提取8月份目标整改问题"):
        # Extract issues to be completed in August
        next_month = (datetime.now().replace(day=1) + pd.DateOffset(months=1)).strftime('%Y-%m')
        df = st.session_state.df
        df['目标设定完成时间'] = pd.to_datetime(df['目标设定完成时间'], errors='coerce')
        upcoming_month_df = df[df['目标设定完成时间'].dt.strftime('%Y-%m') == next_month]

        if not upcoming_month_df.empty:
            st.write("**须在8月份整改完成的整改问题(点击表格右上角下载按钮可直接保存为csv)**")
            st.dataframe(upcoming_month_df)

           