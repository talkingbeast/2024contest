import pandas as pd
import streamlit as st
import plotly.express as px
from io import BytesIO

# Streamlit UI
st.title("任务6：按问题分类自动查询分析")
st.write('请先选择汇总台账文件，然后选择分组类别和问题分类进行分析。')

# Session state for file paths and data
if 'file_path' not in st.session_state:
    st.session_state.file_path = None

if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame()

if 'group_col' not in st.session_state:
    st.session_state.group_col = "所属责任公司"

# File uploader for selecting the summary ledger
uploaded_file = st.file_uploader("选择汇总台账文件", type=['xlsx'])
if uploaded_file:
    st.session_state.file_path = uploaded_file
    st.session_state.df = pd.read_excel(uploaded_file)
    non_numeric_cols = st.session_state.df.select_dtypes(exclude=['number', 'datetime']).columns
    st.session_state.group_col = non_numeric_cols[0] if "所属责任公司" not in non_numeric_cols else "所属责任公司"

# Dropdown for group category
if not st.session_state.df.empty:
    non_numeric_cols = st.session_state.df.select_dtypes(exclude=['number', 'datetime']).columns
    st.session_state.group_col = st.selectbox('选择分组类别:', non_numeric_cols, index=non_numeric_cols.get_loc(st.session_state.group_col))

    # Dropdown for issue category
    issue_categories = st.session_state.df['问题分类'].unique()
    selected_issue = st.selectbox('选择问题分类:', issue_categories)

    # Filtering data based on selected issue category
    filtered_df = st.session_state.df[st.session_state.df['问题分类'] == selected_issue]

    if not filtered_df.empty:
        # Group by selected group column
        grouped = filtered_df.groupby(st.session_state.group_col).agg(
            已整改数量=('是否整改完成', lambda x: (x == '是').sum()),
            未整改数量=('是否整改完成', lambda x: (x == '否').sum()),
            已整改金额=('提示问题金额（万元）', lambda x: filtered_df.loc[x.index, '提示问题金额（万元）'].where(x == '是').sum()),
            未整改金额=('提示问题金额（万元）', lambda x: filtered_df.loc[x.index, '提示问题金额（万元）'].where(x == '否').sum())
        ).reset_index()

        # Create interactive charts
        fig1 = px.bar(grouped, x=st.session_state.group_col, y=['已整改数量', '未整改数量'], title='各公司审计整改统计', labels={'value': '数量', 'variable': '统计类型'})
        fig2 = px.bar(grouped, x=st.session_state.group_col, y=['已整改金额', '未整改金额'], title='各公司审计整改金额统计', labels={'value': '金额（万元）', 'variable': '统计类型'})

        st.plotly_chart(fig1)
        st.plotly_chart(fig2)
        st.write(f'点击表格右上角下载按钮可直接保存为csv')
        st.dataframe(grouped)

        # Save DataFrame to Excel in memory
        def save_df_to_excel(df):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='openpyxl')
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            writer.close()
            output.seek(0)
            return output

        if st.button('下载分析结果'):
            # Save summary to Excel
            excel_bytes = save_df_to_excel(grouped)
            st.download_button(
                label="下载审计整改汇总表格",
                data=excel_bytes,
                file_name='审计整改汇总.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

            # Save charts as HTML files
            fig1_html = fig1.to_html(full_html=False)
            fig2_html = fig2.to_html(full_html=False)

            st.download_button(
                label="下载审计整改数量统计图表",
                data=fig1_html,
                file_name='审计整改数量统计.html',
                mime="text/html"
            )

            st.download_button(
                label="下载审计整改金额统计图表",
                data=fig2_html,
                file_name='审计整改金额统计.html',
                mime="text/html"
            )

else:
    st.info("请选择汇总台账文件。")
