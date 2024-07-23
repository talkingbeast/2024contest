import os
import pandas as pd
from datetime import datetime
import streamlit as st
import tkinter as tk
from tkinter import filedialog
from pathlib import Path

# Function to read all selected xlsx files and combine them into one DataFrame
def read_and_combine_xlsx_files(file_paths):
    combined_df = pd.DataFrame()
    
    for file_path in file_paths:
        df = pd.read_excel(file_path)
        df['来源文件'] = os.path.basename(file_path)  # 添加一列以标识来源文件
        combined_df = pd.concat([combined_df, df], ignore_index=True)
    
    return combined_df

# Function to get the latest ledger file
def get_latest_ledger_file():
    current_dir = Path('.')
    ledger_files = list(current_dir.glob('集团整体审计整改台账_*.xlsx'))
    if not ledger_files:
        return ""
    latest_file = max(ledger_files, key=os.path.getctime)
    return str(latest_file)

# Set up tkinter
root = tk.Tk()
root.withdraw()
root.wm_attributes('-topmost', 1)

# Streamlit UI
st.title("任务2：新增集团所属C公司审计整改台账")
st.write('###### （二）2024年7月，新增集团所属C公司审计整改台账，请将C公司审计整改台账自动汇总至集团整体审计整改台账')
st.write('---')

if 'file_paths' not in st.session_state:
    st.session_state.file_paths = []
if 'original_ledger_path' not in st.session_state:
    st.session_state.original_ledger_path = get_latest_ledger_file()

# File picker button for selecting C公司 files
if st.button('选择新增公司台账文件(支持多选)'):
    file_paths = filedialog.askopenfilenames(master=root, filetypes=[("Excel files", "*.xlsx")])
    if file_paths:
        st.session_state.file_paths = list(file_paths)

if st.session_state.file_paths:
    st.text_input('已选择新增公司台账文件:', ', '.join(st.session_state.file_paths))

# File picker button for selecting the original ledger
if st.button('选择原汇总台账文件'):
    original_ledger_path = filedialog.askopenfilename(master=root, filetypes=[("Excel files", "*.xlsx")])
    if original_ledger_path:
        st.session_state.original_ledger_path = original_ledger_path

if st.session_state.original_ledger_path:
    st.text_input('已选择原汇总台账文件:', st.session_state.original_ledger_path)
else:
    st.info("默认选择程序路径下最新的 '集团整体审计整改台账' 文件。")

if st.session_state.file_paths and st.session_state.original_ledger_path:
    if st.button("重新生成台账"):
        new_data = read_and_combine_xlsx_files(st.session_state.file_paths)
        
        if new_data is not None:
            original_data = pd.read_excel(st.session_state.original_ledger_path, sheet_name='汇总台账')
            combined_data = pd.concat([original_data, new_data], ignore_index=True)

            if '序号' in combined_data.columns:
                combined_data['序号'] = range(1, len(combined_data) + 1)

            current_time = datetime.now().strftime('%Y%m%d%H%M%S')
            output_file = f'集团整体审计整改台账_{current_time}.xlsx'
            
            with pd.ExcelWriter(output_file) as writer:
                combined_data.to_excel(writer, sheet_name='汇总台账', index=False)
            
            st.info(f'新台账已生成并保存为: {output_file}')
            st.write(f'点击表格右上角下载按钮可直接保存为csv')
            st.dataframe(combined_data)
else:
    st.info("请选择新增公司台账文件和原汇总台账文件，然后点击 '重新生成台账' 按钮。")
