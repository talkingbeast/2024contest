import pandas as pd
from datetime import datetime
import streamlit as st
from io import BytesIO

# Initialize session state variables
if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame()

# Function to save DataFrame to Excel in memory
def save_df_to_excel(df, sheet_name='Sheet1'):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

# Streamlit UI
st.title("任务8和任务9：整改事项提取与目标提示")
st.write('请选择汇总台账文件，然后选择功能按钮进行操作。')

# File uploader for selecting the summary ledger
uploaded_file = st.file_uploader("选择汇总台账文件", type=['xlsx'])
if uploaded_file:
    st.session_state.df = pd.read_excel(uploaded_file)
    st.write(f"已加载数据文件: {uploaded_file.name}")

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
            st.write("**本月完成整改事项(点击表格右上角下载按钮可直接保存为xlsx)**")
            
            # Adjust the '序号' column
            if '序号' in current_month_df.columns:
                current_month_df['序号'] = range(1, len(current_month_df) + 1)
            
            st.dataframe(current_month_df)
            
            excel_data = save_df_to_excel(current_month_df, sheet_name=f"本月完成整改事项_{current_month}")
            st.download_button(
                label="下载本月完成整改事项",
                data=excel_data,
                file_name=f"本月完成整改事项_{current_month}.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

    if st.button("提取8月份目标整改问题"):
        # Extract issues to be completed in August
        next_month = (datetime.now().replace(day=1) + pd.DateOffset(months=1)).strftime('%Y-%m')
        df = st.session_state.df
        df['目标设定完成时间'] = pd.to_datetime(df['目标设定完成时间'], errors='coerce')
        upcoming_month_df = df[df['目标设定完成时间'].dt.strftime('%Y-%m') == next_month]

        if not upcoming_month_df.empty:
            st.write("**须在8月份整改完成的整改问题(点击表格右上角下载按钮可直接保存为xlsx)**")
            
            # Adjust the '序号' column
            if '序号' in upcoming_month_df.columns:
                upcoming_month_df['序号'] = range(1, len(upcoming_month_df) + 1)
            
            st.dataframe(upcoming_month_df)
            
            excel_data = save_df_to_excel(upcoming_month_df, sheet_name=f"8月份目标整改问题_{next_month}")
            st.download_button(
                label="下载8月份目标整改问题",
                data=excel_data,
                file_name=f"8月份目标整改问题_{next_month}.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
