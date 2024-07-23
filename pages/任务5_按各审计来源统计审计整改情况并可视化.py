import pandas as pd
import streamlit as st
import plotly.express as px
from io import BytesIO

# Streamlit UI
st.title("任务五：审计整改数据分析")
st.write('请按照要求选择汇总台账文件，然后点击“开始分析”按钮进行数据分析。')

# Initialize session state
if 'file_path' not in st.session_state:
    st.session_state.file_path = None

# File uploader for selecting the summary ledger
uploaded_file = st.file_uploader("选择汇总台账文件", type=['xlsx'])
if uploaded_file:
    st.session_state.file_path = uploaded_file

# Display the selected file path
if st.session_state.file_path:
    st.text_input('已选择汇总台账文件:', st.session_state.file_path.name)

# Function to save DataFrame to Excel in memory
def save_df_to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.close()
    output.seek(0)
    return output

if st.session_state.file_path:
    if st.button('开始分析'):
        # Load data
        df = pd.read_excel(st.session_state.file_path)

        # Ensure required columns exist
        required_columns = [
            '问题来源', '整改计划及措施', '是否整改完成', '提示问题金额（万元）', '验收负责人'
        ]
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"数据中缺少以下必要列: {', '.join(missing_columns)}")
        else:
            # Group by issue source and calculate required metrics
            grouped = df.groupby('问题来源').agg(
                总数量=('整改计划及措施', 'count'),
                已整改数量=('是否整改完成', lambda x: (x == '是').sum()),
                未整改数量=('是否整改完成', lambda x: (x == '否').sum()),
                总金额=('提示问题金额（万元）', 'sum'),
                已整改金额=('提示问题金额（万元）', lambda x: df.loc[x.index][df['是否整改完成'] == '是']['提示问题金额（万元）'].sum()),
                未整改金额=('提示问题金额（万元）', lambda x: df.loc[x.index][df['是否整改完成'] == '否']['提示问题金额（万元）'].sum())
            )
            
            # Add acceptance personnel information
            acceptance_personnel = df.groupby('问题来源')['验收负责人'].apply(lambda x: ', '.join(x.dropna().unique()))
            grouped = grouped.join(acceptance_personnel)

            grouped.reset_index(inplace=True)

            # Save summary to Excel
            excel_bytes = save_df_to_excel(grouped)
            st.download_button(
                label="下载审计整改汇总表格",
                data=excel_bytes,
                file_name='审计整改汇总.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

            st.write(f'点击表格右上角下载按钮可直接保存为csv')
            st.dataframe(grouped)

            # Create interactive charts
            fig1 = px.bar(grouped, x='问题来源', y=['已整改数量', '未整改数量'],
                          title='各审计来源审计整改数量统计',
                          labels={'value': '数量', 'variable': '统计类型'})
            st.plotly_chart(fig1)

            fig2 = px.bar(grouped, x='问题来源', y=['已整改金额', '未整改金额'],
                          title='各审计来源审计整改金额统计',
                          labels={'value': '金额（万元）', 'variable': '统计类型'})
            st.plotly_chart(fig2)

            # Save charts as HTML files
            fig1_html = fig1.to_html(full_html=False)
            fig2_html = fig2.to_html(full_html=False)

            st.download_button(
                label="下载各审计来源审计整改数量统计图表",
                data=fig1_html,
                file_name='各审计来源审计整改数量统计.html',
                mime="text/html"
            )

            st.download_button(
                label="下载各审计来源审计整改金额统计图表",
                data=fig2_html,
                file_name='各审计来源审计整改金额统计.html',
                mime="text/html"
            )

else:
    st.info("请选择汇总台账文件，然后点击 '开始分析' 按钮进行分析。")
