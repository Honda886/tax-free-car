from docx import Document
import pandas as pd
import streamlit as st
import numpy as np
import os

# 设置网页信息 
st.set_page_config(page_title="工信部-减免车辆购置税的新能源汽车车型目录-汇总一览", page_icon=":racing_car:", layout="wide")

# 识别并提取Word文档中的表格
@st.cache_data
def extract_tables_from_docx(docx_path):
    doc = Document(docx_path)
    tables = []

    for table in doc.tables:
        data = []
        for row in table.rows:
            row_data = []

            for cell in row.cells:
                row_data.append(cell.text)
                
            data.append(row_data)

        df = pd.DataFrame(data)
        tables.append(df)

    return tables


# 遍历文件夹中的所有Word文档
@st.cache_data
def extract_data_from_folder(folder_path):

    all_tables = []

    for filename in os.listdir(folder_path):

        if filename.endswith('.docx'):
            docx_path = os.path.join(folder_path, filename)
            tables = extract_tables_from_docx(docx_path)
            all_tables.extend(tables)

    return all_tables


# 将所有提取出来的表格数据合并到一个大的 DataFrame 中，便于后续的数据处理与分析。
@st.cache_data
def integrate_data(tables_list):
    all_data = pd.concat(tables_list, ignore_index=True)
    
    return all_data

# 数据清洗
@st.cache_data
def clean_data(unique_df):
    unique_df = integrated_data.drop_duplicates(keep='first')
    
    return unique_df

# 将以上所有步骤合并，可以创建一个脚本来实现从提取到分析的完整流程。
folder_path = './06dmi/'
tables_list = extract_data_from_folder(folder_path)
integrated_data = integrate_data(tables_list)
cleaned_data = clean_data(integrated_data)

# 交由streamlit呈现web表格

# 页面标题
st.header("工信部-减免车辆购置税的新能源汽车车型目录-汇总一览-更新日期：2024-09-11")

# 创建一个新列 '搜索字段'，将所有列合并为一个字符串，便于快速搜索
cleaned_data['搜索字段'] = cleaned_data.astype(str).apply(' '.join, axis=1).str.lower()

# 定义一个 session state 来存储搜索框的值和初始数据
if 'search_text' not in st.session_state:
    st.session_state.search_text = ""

# 搜索框
search_input = st.text_input(
    label="🔍 搜索你中意的爱车吧！按下 Enter 键确认", 
    value=st.session_state.search_text,  # 初始值为 session_state 中的搜索值
    placeholder="请输入搜索关键字...",
)

# 重置按钮
if st.button("（↻）重置到初始状态"):
    # 点击重置按钮时，清空搜索框并返回初始状态
    st.session_state.search_text = ""  # 清空搜索框
    search_input = ""  # 也将当前搜索框输入清空

# 如果搜索框有输入内容，过滤表格数据
if search_input:
    # 将用户输入的值保存到 session_state
    st.session_state.search_text = search_input
    
    # 只搜索预处理过的 '搜索字段' 列
    filtered_data = cleaned_data[cleaned_data['搜索字段'].str.contains(search_input.lower())]
    
    # 展示过滤后的数据，不包含 '搜索字段' 列
    st.dataframe(filtered_data.drop(columns=['搜索字段']))
else:
    # 否则展示初始完整数据，不包含 '搜索字段' 列
    st.dataframe(cleaned_data.drop(columns=['搜索字段']))