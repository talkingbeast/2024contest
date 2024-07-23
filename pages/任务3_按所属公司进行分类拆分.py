import os
import pandas as pd
from datetime import datetime
import streamlit as st
import zipfile
from io import BytesIO
from pathlib import Path

# 获取最新的台账文件
def get_latest_ledger_file():
    current_dir = Path('.')
    ledger_files = list(current_dir.glob('集团整体审计整改台账_*.xlsx'))
    if not ledger_files:
        return None
    latest_file = max(ledger_files, key=os.path.getctime)
    return str(latest_file)

# 按“所属责任公司”字段拆分数据，并保存到不同的工作簿中
def split_and_save_by_company(original_ledger_file):
    # 读取原汇总台账文件
    original_data = pd.read_excel(original_ledger_file, sheet_name=None)
    original_data = pd.concat(original_data.values(), ignore_index=True)

    # 按“所属责任公司”字段分组
    grouped = original_data.groupby('所属责任公司')
    
    output_folder = "拆分台账"
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
    
    return output_folder

# 打包文件夹
def zip_folder(folder_path):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                zipf.write(file_path, os.path.relpath(file_path, folder_path))
    zip_buffer.seek(0)
    return zip_buffer

# Streamlit UI
st.title("任务3：按所属公司拆分审计整改台账")
st.write('###### 将合并的审计整改台账按公司拆分成独立工作簿，并在每个工作簿的最后一行添加提示信息。')
st.write('---')

if 'original_ledger_path' not in st.session_state:
    st.session_state.original_ledger_path = get_latest_ledger_file()

# 上传原汇总台账文件
original_ledger_file = st.file_uploader("选择原汇总台账文件", type=['xlsx'])

if original_ledger_file:
    if st.button("生成拆分台账"):
        # 生成按公司拆分后的台账文件
        output_folder = split_and_save_by_company(original_ledger_file)
        
        # 打包拆分后的文件
        zip_buffer = zip_folder(output_folder)
        
        st.success(f"拆分台账已生成并保存在文件夹：{output_folder}")
        
        # 提供下载链接
        st.download_button(
            label="下载拆分台账压缩包",
            data=zip_buffer,
            file_name=f'拆分台账_{datetime.now().strftime("%Y%m%d%H%M%S")}.zip',
            mime="application/zip"
        )
else:
    st.info("请上传原汇总台账文件，然后点击 '生成拆分台账' 按钮。")
