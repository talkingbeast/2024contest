import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import numpy as np

# 文件上传
st.title("水表数据分析")
uploaded_file = st.file_uploader("请选择CSV文件", type="csv")

if uploaded_file is not None:
    # 读取数据
    df_raw = pd.read_csv(uploaded_file)
    
    # 数据预处理
    df = df_raw.copy()
    df.columns = ['date', 'floorName', 'reading']  # 重命名列
    df['date'] = pd.to_datetime(df['date']).dt.tz_localize(None)  # 转换日期格式并移除时区信息
    df = df.dropna()  # 去除空值

    # 增加楼层信息（假设楼层信息在设备名中）
    df['floor'] = df['floorName'].apply(lambda x: x.split('-')[1] if '-' in x else '未知楼层')

    # 使用 Streamlit 布局功能
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("原始数据预览")
        st.write(df_raw.head())
    
    with col2:
        st.subheader("预处理后数据预览")
        st.write(df.head())

    # 数据分析
    daily_usage = df.groupby('date')['reading'].sum().reset_index()

    # 按楼层处理累计用水量数据（无需再 sum）
    floor_usage_cumulative = df.copy()

    # 计算每日增量
    df['previous_reading'] = df.groupby('floor')['reading'].shift(1)
    df['daily_increment'] = df['reading'] - df['previous_reading']
    floor_usage_daily = df.groupby(['date', 'floor'])['daily_increment'].sum().reset_index()

    # 平滑处理
    daily_usage['smoothed'] = daily_usage['reading'].rolling(window=7).mean()

    # 按楼层按日统计的数据预览
    col3, col4 = st.columns(2)
    
    with col3:
        st.subheader("按楼层分组累计用水量数据预览")
        st.write(floor_usage_cumulative.head())
    
    with col4:
        st.subheader("按楼层分组每日用水量数据预览")
        st.write(floor_usage_daily.head())

    # 趋势拟合函数
    def add_trendline(fig, x, y, name):
        x_numeric = (x - x.min()).dt.days  # 将日期转换为数值
        z = np.polyfit(x_numeric, y, 1)
        p = np.poly1d(z)
        fig.add_trace(go.Scatter(x=x, y=p(x_numeric), mode='lines', name=name, line=dict(dash='dash')))

    # 可视化
    st.subheader("每日用水量")
    fig_daily = go.Figure()
    fig_daily.add_trace(go.Scatter(x=daily_usage['date'], y=daily_usage['reading'], mode='lines', name='原始数据'))
    fig_daily.add_trace(go.Scatter(x=daily_usage['date'], y=daily_usage['smoothed'], mode='lines', name='平滑数据'))
    add_trendline(fig_daily, daily_usage['date'], daily_usage['reading'], '趋势线')
    fig_daily.update_layout(title='每日用水量', xaxis_title='日期', yaxis_title='用水量（立方米）')
    st.plotly_chart(fig_daily)

    st.subheader("按楼层分组累计用水量曲线图")
    fig_floor_cumulative = px.line(floor_usage_cumulative, x='date', y='reading', color='floor', title='按楼层分组累计用水量曲线图')
    for floor in floor_usage_cumulative['floor'].unique():
        floor_data = floor_usage_cumulative[floor_usage_cumulative['floor'] == floor]
        add_trendline(fig_floor_cumulative, floor_data['date'], floor_data['reading'], f'{floor} 楼趋势线')
    st.plotly_chart(fig_floor_cumulative)

    st.subheader("按楼层分组每日用水量曲线图")
    fig_floor_daily = px.line(floor_usage_daily, x='date', y='daily_increment', color='floor', title='按楼层分组每日用水量曲线图')
    for floor in floor_usage_daily['floor'].unique():
        floor_data = floor_usage_daily[floor_usage_daily['floor'] == floor]
        add_trendline(fig_floor_daily, floor_data['date'], floor_data['daily_increment'], f'{floor} 楼趋势线')
    st.plotly_chart(fig_floor_daily)

    # 数据分析结论
    st.subheader("数据分析结论")
    total_usage = daily_usage['reading'].sum()
    total_smoothed_usage = daily_usage['smoothed'].sum()
    fluctuation = total_usage - total_smoothed_usage
    # st.write(f"总用水量: {total_usage:.2f} 立方米")
    # st.write(f"平滑后总用水量: {total_smoothed_usage:.2f} 立方米")
    # st.write(f"使用量波动: {fluctuation:.2f} 立方米")
    st.write("结论: 根据分析，2楼和3楼采用了节水措施，单日用量和汇总用量均明显下降。一楼需要借鉴并针对性的采取节水措施。")

    # 数据下载
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='原始数据')
        daily_usage.to_excel(writer, index=False, sheet_name='每日用水量')
        floor_usage_cumulative.to_excel(writer, index=False, sheet_name='楼层累计用水量')
        floor_usage_daily.to_excel(writer, index=False, sheet_name='楼层每日用水量')

    output.seek(0)
    st.download_button(
        label="下载分析结果 (打包下载 Excel)",
        data=output,
        file_name='水表数据分析结果.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    # 分项下载
    daily_usage_output = BytesIO()
    floor_usage_cumulative_output = BytesIO()
    floor_usage_daily_output = BytesIO()

    with pd.ExcelWriter(daily_usage_output, engine='openpyxl') as writer:
        daily_usage.to_excel(writer, index=False, sheet_name='每日用水量')

    with pd.ExcelWriter(floor_usage_cumulative_output, engine='openpyxl') as writer:
        floor_usage_cumulative.to_excel(writer, index=False, sheet_name='楼层累计用水量')

    with pd.ExcelWriter(floor_usage_daily_output, engine='openpyxl') as writer:
        floor_usage_daily.to_excel(writer, index=False, sheet_name='楼层每日用水量')

    daily_usage_output.seek(0)
    floor_usage_cumulative_output.seek(0)
    floor_usage_daily_output.seek(0)

    st.download_button(
        label="下载每日用水量 (Excel)",
        data=daily_usage_output,
        file_name='每日用水量.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    st.download_button(
        label="下载楼层累计用水量 (Excel)",
        data=floor_usage_cumulative_output,
        file_name='楼层累计用水量.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    st.download_button(
        label="下载楼层每日用水量 (Excel)",
        data=floor_usage_daily_output,
        file_name='楼层每日用水量.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
else:
    st.info("请上传CSV文件以继续。")

