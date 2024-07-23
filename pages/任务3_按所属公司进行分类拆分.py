import os
import pandas as pd
from datetime import datetime
import streamlit as st
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import zipfile

# 读取指定路径下的所有xlsx文件
def read_all_xlsx_files(file_paths):
    data_frames = []
    
    for file_path in file_paths:
        df = pd.read_excel(file_path)
        data_frames.append({'file_name': os.path.basename(file_path), 'df': df})
    
    return data_frames

# 获取最新的台账文件
def get_latest_ledger_file():
    current_dir = Path('.')
    ledger_files = list(current_dir.glob('集团整体审计整改台账_*.xlsx'))
    if not ledger_files:
        return ""
    latest_file = max(ledger_files, key=os.path.getctime)
    return str(latest_file)

# 按“所属责任公司”字段拆分数据，并保存到不同的工作簿中
def split_and_save_by_company(data_frames, output_folder):
    # 合并所有数据
    all_data = pd.concat([item['df'] for item in data_frames], ignore_index=True)
    
    # 按“所属责任公司”字段分组
    grouped = all_data.groupby('所属责任公司')
    
    # 确保输出文件夹存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    for company, group in grouped:
        # 重新编号序号
        if '序号' in group.columns:
            group['序号'] = range(1, len(group) + 1)
        
        # 生成每个公司的工作簿文件名
        company_file_name = os.path.join(output_folder, f'{company}.xlsx')
        
        # 将每个公司的数据保存到一个工作簿中
        with pd.ExcelWriter(company_file_name, engine='openpyxl') as writer:
            group.to_excel(writer, sheet_name='整改台账', index=False)
            
            # 在每个工作表的最后一行添加提示信息
            sheet = writer.sheets['整改台账']
            row_num = len(group) + 2  # 计算提示信息的位置
            sheet.cell(row=row_num, column=1, value="注意：请按规则进行填报")

# 打包文件夹
def zip_folder(folder_path, output_zip):
    with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                zipf.write(file_path, os.path.relpath(file_path, folder_path))

# 设置Tkinter
root = tk.Tk()
root.withdraw()
root.wm_attributes('-topmost', 1)

# Streamlit UI
st.title("任务3：按所属公司拆分审计整改台账")
st.write('###### 将合并的审计整改台账按公司拆分成独立工作簿，并在每个工作簿的最后一行添加提示信息。')
st.write('---')

if 'file_paths' not in st.session_state:
    st.session_state.file_paths = []
if 'original_ledger_path' not in st.session_state:
    st.session_state.original_ledger_path = get_latest_ledger_file()
if 'output_folder' not in st.session_state:
    st.session_state.output_folder = ""

# 选择原汇总台账文件
if st.button('选择原汇总台账文件'):
    original_ledger_path = filedialog.askopenfilename(master=root, filetypes=[("Excel files", "*.xlsx")])
    if original_ledger_path:
        st.session_state.original_ledger_path = original_ledger_path

if st.session_state.original_ledger_path:
    st.text_input('已选择原汇总台账文件:', st.session_state.original_ledger_path)
else:
    st.info("默认选择程序路径下最新的 '集团整体审计整改台账' 文件。")

# 选择输出文件夹
if st.button('选择保存文件夹'):
    output_folder = filedialog.askdirectory(master=root)
    if output_folder:
        st.session_state.output_folder = output_folder
    else:
        st.warning("未选择文件夹，无法保存拆分后的台账。")

if 'output_folder' in st.session_state and st.session_state.original_ledger_path:
    if st.button("生成拆分台账"):
        # 读取原汇总台账文件
        original_ledger_path = st.session_state.original_ledger_path
        
        with pd.ExcelFile(original_ledger_path) as xls:
            original_data = pd.concat(pd.read_excel(xls, sheet_name=None), ignore_index=True)
        
        # 读取所有台账文件并将其合并
        data_frames = [{'file_name': os.path.basename(original_ledger_path), 'df': original_data}]
        
        # 生成按公司拆分后的台账文件
        output_folder = st.session_state.output_folder
        split_and_save_by_company(data_frames, output_folder)
        
        # 打包拆分后的文件
        zip_file_path = Path('.') / f'拆分台账_{datetime.now().strftime("%Y%m%d%H%M%S")}.zip'
        zip_folder(output_folder, zip_file_path)
        
        st.success(f"拆分台账已生成并保存到: {output_folder}")
        
        # 提供下载链接
        with open(zip_file_path, 'rb') as f:
            st.download_button(label="下载拆分台账压缩包", data=f, file_name=zip_file_path.name)

else:
    if not 'output_folder' in st.session_state:
        st.info("请选择保存文件夹，然后点击 '生成拆分台账' 按钮。")
    if not st.session_state.original_ledger_path:
        st.info("请选择原汇总台账文件，然后点击 '生成拆分台账' 按钮。")
