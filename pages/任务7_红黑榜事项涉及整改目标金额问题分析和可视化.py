import pandas as pd
import streamlit as st
import tkinter as tk
from tkinter import filedialog
import plotly.express as px
import os
from pathlib import Path

# Set up tkinter
root = tk.Tk()
root.withdraw()
root.wm_attributes('-topmost', 1)

# Streamlit UI
st.title("任务7：红黑榜事项整改目标金额统计")
st.write('请先选择汇总台账文件，然后进行分析。')
st.write('原始数据有问题，没有2024上半年已整改金额字段，只能使用 本年度金额类累计整改进展 字段，请知悉！')

# Session state for file paths and data
if 'file_path' not in st.session_state:
    st.session_state.file_path = None

if 'save_folder' not in st.session_state:
    st.session_state.save_folder = ""

if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame()

# Function to select file
def select_file():
    file_path = filedialog.askopenfilename(master=root, filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        st.session_state.file_path = file_path
        st.session_state.df = pd.read_excel(file_path)

# Function to select folder
def select_folder():
    folder = filedialog.askdirectory(master=root)
    if folder:
        st.session_state.save_folder = folder

# File picker for selecting the summary ledger
if st.button('选择汇总台账文件'):
    select_file()

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

    if st.button('保存分析结果'):
        select_folder()
        if st.session_state.save_folder:
            base_path = Path(st.session_state.save_folder)
            
            # Save 2024上半年数据
            path_1 = base_path / '2024上半年整改目标金额_vs_已整改金额.xlsx'
            with pd.ExcelWriter(path_1) as writer:
                summary_df[['指标', '金额（万元）']].iloc[:2].to_excel(writer, sheet_name='分析结果', index=False)
            fig1.write_html(base_path / '2024上半年整改目标金额_vs_已整改金额.html')

            # Save 2024全年数据
            path_2 = base_path / '2024全年整改目标金额_vs_已整改金额.xlsx'
            with pd.ExcelWriter(path_2) as writer:
                summary_df[['指标', '金额（万元）']].iloc[2:4].to_excel(writer, sheet_name='分析结果', index=False)
            fig2.write_html(base_path / '2024全年整改目标金额_vs_已整改金额.html')

            # Save 总体发现问题数据
            path_3 = base_path / '总体发现问题提示金额_vs_已整改金额.xlsx'
            with pd.ExcelWriter(path_3) as writer:
                summary_df[['指标', '金额（万元）']].iloc[4:].to_excel(writer, sheet_name='分析结果', index=False)
            fig3.write_html(base_path / '总体发现问题提示金额_vs_已整改金额.html')

            st.success(f"汇总表格和图表已保存至: {st.session_state.save_folder}")

else:
    st.info("请选择汇总台账文件。")
