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

# 添加标题
st.header("工信部-减免车辆购置税的新能源汽车车型目录-汇总一览-更新日期：2024-09-11")

# 添加搜索框和按钮
search_term = st.text_input(label="🔍 搜索你中意的爱车吧！", placeholder="请输入搜索关键字...点击搜索按钮确认",)
search_button = st.button("搜索")
reset_button = st.button("（↻）重置到初始状态")

# 初始化一个变量来存储过滤后的数据
filtered_data = cleaned_data.copy()

# 如果点击了搜索按钮
if search_button:
    # 使用矢量化方法进行字符串匹配，提高速度
    def contains_search_term(row):
        return any(search_term.lower() in str(cell).lower() for cell in row)
    filtered_data = cleaned_data[cleaned_data.apply(contains_search_term, axis=1)]

# 如果点击了重置按钮
if reset_button:
    # 重置数据并清空搜索框
    filtered_data = cleaned_data
    search_term = ""

# 显示过滤后的数据
st.dataframe(filtered_data)