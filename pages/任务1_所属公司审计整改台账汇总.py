import os
import pandas as pd
from datetime import datetime
import streamlit as st
import tkinter as tk
from tkinter import filedialog

# Function to read all xlsx files from the folder and combine them into one DataFrame
def read_and_combine_xlsx_files(folder_path):
    all_files = os.listdir(folder_path)
    xlsx_files = [file for file in all_files if file.endswith('.xlsx')]
    
    if not xlsx_files:
        st.error("选择的文件夹中没有找到任何 xlsx 文件。")
        return None
    
    combined_df = pd.DataFrame()
    
    for file in xlsx_files:
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path)
        df['来源文件'] = file  # 添加一列以标识来源文件
        combined_df = pd.concat([combined_df, df], ignore_index=True)
    
    return combined_df

# Set up tkinter
root = tk.Tk()
root.withdraw()
root.wm_attributes('-topmost', 1)

# Streamlit UI
st.title("任务1：所属公司审计整改台账汇总")

st.write("请点击 '选择文件夹' 按钮选择台账文件文件夹，然后点击 '台账汇总' 按钮，程序将遍历合并为汇总台账。")

if 'folder_path' not in st.session_state:
    st.session_state.folder_path = ""

# Folder picker button
if st.button('选择文件夹'):
    folder_path = filedialog.askdirectory(master=root)
    if folder_path:
        st.session_state.folder_path = folder_path

if st.session_state.folder_path:
    st.text_input('已选择文件夹:', st.session_state.folder_path)

    if not os.path.isdir(st.session_state.folder_path):
        st.error("输入的路径不是有效的文件夹。")
    else:
        # Button to start the summary process
        if st.button("台账汇总"):
            combined_df = read_and_combine_xlsx_files(st.session_state.folder_path)
                # 重新编号序号
            if '序号' in combined_df.columns:
                combined_df['序号'] = range(1, len(combined_df) + 1)
            
            if combined_df is not None:
                current_time = datetime.now().strftime('%Y%m%d%H%M%S')
                output_file = f'集团整体审计整改台账_{current_time}.xlsx'
                
                with pd.ExcelWriter(output_file) as writer:
                    combined_df.to_excel(writer, sheet_name='汇总台账', index=False)
                
                st.info(f'文件已保存为: {output_file}')
                st.write(f'点击表格右上角下载按钮可直接保存为csv')
                st.dataframe(combined_df)
else:
    st.info("请选择一个文件夹以继续。")


