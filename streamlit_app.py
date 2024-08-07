import streamlit as st
import os
import zipfile
from io import BytesIO

st.title('第七届职工职业技能竞赛大数据处理与审计决赛题目')
st.header('—审计整改台账分析')
st.write('###### by 智慧能源-18503')
st.write('---')

# 数据文件路径
data_dir = '决赛题目/决赛数据'
files_to_download = [
    '口腔养护消费者洞察.pdf',
    '处理后pdf_口腔养护消费者洞察.pdf',
    '水表数据-2019-09-28.csv'
]

# 创建一个zip文件
def create_zip_file(file_paths):
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, 'w') as z:
        for file_path in file_paths:
            z.write(file_path, os.path.basename(file_path))
    buffer.seek(0)
    return buffer

# 检查文件是否存在，并生成完整路径
file_paths = [os.path.join(data_dir, file) for file in files_to_download]
existing_files = [file for file in file_paths if os.path.exists(file)]

# 说明下载的文件结构
st.write('下载的文件结构:')
st.write('```\n决赛数据.zip\n├── 口腔养护消费者洞察.pdf\n├── 处理后pdf_口腔养护消费者洞察.pdf\n└── 水表数据-2019-09-28.csv\n```')

# 创建zip文件并提供下载
if existing_files:
    zip_buffer = create_zip_file(existing_files)
    st.download_button(
        label='下载测试数据',
        data=zip_buffer,
        file_name='决赛数据.zip',
        mime='application/zip'
    )
else:
    st.error('文件未找到或不存在')

st.write('---')
