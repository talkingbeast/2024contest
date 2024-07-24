import streamlit as st
import os
import zipfile
from io import BytesIO

st.title('第七届职工职业技能竞赛大数据处理与审计初赛题目')
st.header('—审计整改台账分析')
st.write('###### by 智慧能源-18503')
st.write('---')

# st.sidebar.title('导航')

# 指定数据包文件夹路径
data_package_folder = '数据包'

# 手动构建文件夹结构
folder_structure = """
数据包
│  C公司.xlsx
│
└─2024年6月七彩集团各所属公司审计整改台账
        A公司.xlsx
        B公司.xlsx
        创业公司.xlsx
        压力公司.xlsx
        境外公司.xlsx
        江东公司.xlsx
        蜀国公司.xlsx
        魏国公司.xlsx
"""

# 显示文件夹结构
st.write("请点击以下按钮下载数据包文件，以便测试页面的各个功能。")
st.write("下面是数据包文件的结构：")
st.code(folder_structure, language='plaintext')

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

# 创建并提供下载链接
def create_zip_and_download(folder_path, zip_filename):
    zip_buffer = zip_folder(folder_path)
    st.download_button(
        label="下载数据包文件",
        data=zip_buffer,
        file_name=zip_filename,
        mime='application/zip'
    )

# 创建并提供下载链接
create_zip_and_download(data_package_folder, 'data_package.zip')
