import pandas as pd
import streamlit as st
import plotly.express as px
from io import BytesIO

# Streamlit UI
st.title("任务7：红黑榜事项整改目标金额统计")
st.write('请先选择汇总台账文件，然后进行分析。')
st.write('原始数据有问题，没有2024上半年已整改金额字段，只能使用 本年度金额类累计整改进展 字段，请知悉！')

# Session state for file paths and data
if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame()

# File uploader for selecting the summary ledger
uploaded_file = st.file_uploader("选择汇总台账文件", type=['xlsx'])
if uploaded_file:
    st.session_state.df = pd.read_excel(uploaded_file)

if not st.session_state.df.empty:
    # Convert numeric columns to appropriate data types
    numeric_columns = [
        '提示问题金额（万元）', '2024年上半年度金额类整改目标', '本年度金额类累计整改进展',
        '2024年年度金额类整改目标', '金额类累计整改进展'
    ]
    for col in numeric_columns:
        st.session_state.df[col] = pd.to_numeric(st.session_state.df[col], errors='coerce').fillna(0)

    # Filtering data for red and black list items
    red_black_df = st.session_state.df[st.session_state.df['红黑榜事项']=='是']

    # Aggregate data
    summary_data = {
        '2024上半年整改目标金额': red_black_df['2024年上半年度金额类整改目标'].sum(),
        '2024上半年已整改金额': red_black_df['本年度金额类累计整改进展'].sum(),
        '2024全年整改目标金额': red_black_df['2024年年度金额类整改目标'].sum(),
        '2024全年已整改金额': red_black_df['本年度金额类累计整改进展'].sum(),
        '总体发现问题提示金额': red_black_df['提示问题金额（万元）'].sum(),
        '总体已整改金额': red_black_df['金额类累计整改进展'].sum(),
    }

    # Display summary data in a table
    summary_df = pd.DataFrame(list(summary_data.items()), columns=['指标', '金额（万元）'])
    st.write("**分析结果(点击表格右上角下载按钮可直接保存为csv)**")
    st.dataframe(summary_df)

    # Create interactive charts
    fig1 = px.bar(
        x=['2024上半年整改目标金额', '2024上半年已整改金额'],
        y=[summary_data['2024上半年整改目标金额'], summary_data['2024上半年已整改金额']],
        title='2024上半年整改目标金额 vs 已整改金额',
        labels={'x': '类别', 'y': '金额（万元）'}
    )

    fig2 = px.bar(
        x=['2024全年整改目标金额', '2024全年已整改金额'],
        y=[summary_data['2024全年整改目标金额'], summary_data['2024全年已整改金额']],
        title='2024全年整改目标金额 vs 已整改金额',
        labels={'x': '类别', 'y': '金额（万元）'}
    )

    fig3 = px.bar(
        x=['总体发现问题提示金额', '总体已整改金额'],
        y=[summary_data['总体发现问题提示金额'], summary_data['总体已整改金额']],
        title='总体发现问题提示金额 vs 已整改金额',
        labels={'x': '类别', 'y': '金额（万元）'}
    )

    st.plotly_chart(fig1)
    st.plotly_chart(fig2)
    st.plotly_chart(fig3)

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
        excel_bytes = save_df_to_excel(summary_df)
        st.download_button(
            label="下载审计整改汇总表格",
            data=excel_bytes,
            file_name='审计整改汇总.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

        # Save charts as HTML files
        fig1_html = fig1.to_html(full_html=False)
        fig2_html = fig2.to_html(full_html=False)
        fig3_html = fig3.to_html(full_html=False)

        st.download_button(
            label="下载2024上半年整改目标金额 vs 已整改金额图表",
            data=fig1_html,
            file_name='2024上半年整改目标金额_vs_已整改金额.html',
            mime="text/html"
        )

        st.download_button(
            label="下载2024全年整改目标金额 vs 已整改金额图表",
            data=fig2_html,
            file_name='2024全年整改目标金额_vs_已整改金额.html',
            mime="text/html"
        )

        st.download_button(
            label="下载总体发现问题提示金额 vs 已整改金额图表",
            data=fig3_html,
            file_name='总体发现问题提示金额_vs_已整改金额.html',
            mime="text/html"
        )

else:
    st.info("请选择汇总台账文件。")
