import streamlit as st

st.title('第七届职工职业技能竞赛大数据处理与审计初赛题目')
st.header('—审计整改台账分析')
st.write('---')
st.write('###### 云南智慧能源股份有限公司 付仕蛟')
st.sidebar.title('导航')
page = st.sidebar.selectbox('选择页面', ['任务1：所属公司审计整改台账汇总'])

if page == '任务1：所属公司审计整改台账汇总':
    import pages.任务1_所属公司审计整改台账汇总
