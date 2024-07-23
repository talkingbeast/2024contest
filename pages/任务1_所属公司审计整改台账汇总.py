import os
import pandas as pd
from datetime import datetime
import streamlit as st

# 函数：读取所有上传的 xlsx 文件并将它们合并为一个 DataFrame
def read_and_combine_xlsx_files(uploaded_files):
    combined_df = pd.DataFrame()

    for uploaded_file in uploaded_files:
        df = pd.read_excel(uploaded_file)
        df['来源文件'] = uploaded_file.name  # 添加一列以标识来源文件
        combined_df = pd.concat([combined_df, df], ignore_index=True)
    
    return combined_df

# Streamlit UI
st.title("任务1：所属公司审计整改台账汇总")

st.write("请上传台账文件，然后点击 '台账汇总' 按钮，程序将遍历合并为汇总台账。")

# 文件上传小部件
uploaded_files = st.file_uploader("选择多个文件", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    st.write("已选择文件:")
    for file in uploaded_files:
        st.write(file.name)

    # 点击按钮开始汇总处理
    if st.button("台账汇总"):
        combined_df = read_and_combine_xlsx_files(uploaded_files)
        # 重新编号序号
        if '序号' in combined_df.columns:
            combined_df['序号'] = range(1, len(combined_df) + 1)
        
        if combined_df is not None:
            current_time = datetime.now().strftime('%Y%m%d%H%M%S')
            output_file = f'集团整体审计整改台账_{current_time}.xlsx'
            
            # 保存汇总文件
            with pd.ExcelWriter(output_file) as writer:
                combined_df.to_excel(writer, sheet_name='汇总台账', index=False)
            
            st.info(f'文件已保存为: {output_file}')
            st.write(f'点击表格右上角下载按钮可直接保存为csv')
            st.dataframe(combined_df)
else:
    st.info("请选择上传文件以继续。")
