import pandas as pd
from datetime import datetime
import streamlit as st
from io import BytesIO

# 函数：读取所有选择的 xlsx 文件并将它们合并为一个 DataFrame
def read_and_combine_xlsx_files(uploaded_files):
    combined_df = pd.DataFrame()
    
    for uploaded_file in uploaded_files:
        df = pd.read_excel(uploaded_file)
        df['来源文件'] = uploaded_file.name  # 添加一列以标识来源文件
        combined_df = pd.concat([combined_df, df], ignore_index=True)
    
    return combined_df

# 函数：获取最新的台账文件
def get_latest_ledger_file():
    # 假设这里提供了路径的代码，你需要自己实现这个部分
    # 因为在 Streamlit Cloud 上不能直接访问文件系统
    return ""

# Streamlit UI
st.title("任务2：新增集团所属C公司审计整改台账")
st.write('###### （二）2024年7月，新增集团所属C公司审计整改台账，请将C公司审计整改台账自动汇总至集团整体审计整改台账')
st.write('---')

if 'file_paths' not in st.session_state:
    st.session_state.file_paths = []
if 'original_ledger_file' not in st.session_state:
    st.session_state.original_ledger_file = None

# 文件上传小部件：选择新增公司台账文件
uploaded_files = st.file_uploader("选择新增公司台账文件 (支持多选)", type=['xlsx'], accept_multiple_files=True)
if uploaded_files:
    st.session_state.file_paths = uploaded_files
    st.text_input('已选择新增公司台账文件:', ', '.join([file.name for file in uploaded_files]))

# 文件上传小部件：选择原汇总台账文件
original_ledger_file = st.file_uploader("选择原汇总台账文件", type=['xlsx'])
if original_ledger_file:
    st.session_state.original_ledger_file = original_ledger_file
    st.text_input('已选择原汇总台账文件:', original_ledger_file.name)

if st.session_state.file_paths and st.session_state.original_ledger_file:
    if st.button("重新生成台账"):
        new_data = read_and_combine_xlsx_files(st.session_state.file_paths)
        
        if new_data is not None:
            original_data = pd.read_excel(st.session_state.original_ledger_file, sheet_name='汇总台账')
            combined_data = pd.concat([original_data, new_data], ignore_index=True)

            if '序号' in combined_data.columns:
                combined_data['序号'] = range(1, len(combined_data) + 1)

            current_time = datetime.now().strftime('%Y%m%d%H%M%S')
            output_file_name = f'集团整体审计整改台账_{current_time}.xlsx'
            st.info(f"台账已生成，文件名为{output_file_name}")
            
            # 将 DataFrame 保存到 BytesIO 中，以便于下载
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                combined_data.to_excel(writer, sheet_name='汇总台账', index=False)
            output.seek(0)  # 重置 BytesIO 对象的指针位置
            
            # 提供下载按钮
            st.download_button(
                label="下载生成的台账",
                data=output,
                file_name=output_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("请选择新增公司台账文件和原汇总台账文件，然后点击 '重新生成台账' 按钮。")
