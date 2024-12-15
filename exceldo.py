import streamlit as st
import pandas as pd
import os
import zipfile
from datetime import datetime
import shutil

# 页面标题
st.title("Excel文件处理工具")

# 上传文件
uploaded_files = st.file_uploader("请上传Excel文件", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    # 创建以当前时间命名的文件夹
    current_time = datetime.now().strftime("%Y%m%d%H%M%S")
    new_folder_name = f"processed_files_{current_time}"
    os.makedirs(new_folder_name, exist_ok=True)

    # 读取每个Excel文件的C2单元格内容并重命名文件
    for uploaded_file in uploaded_files:
        # 读取Excel文件
        try:
            df = pd.read_excel(uploaded_file, header=None)  # 不使用第一行作为标题
        except Exception as e:
            st.error(f"无法读取文件 {uploaded_file.name}: {e}")
            continue  # 跳过无法读取的文件
        
        # 获取C2单元格内容
        try:
            # 确保检查C2单元格是否存在，并处理可能的空值
            c2_content = df.iloc[1, 2] if df.shape[0] > 1 and df.shape[1] > 2 else "Unknown"
            
            # 对内容进行额外的验证和清理
            if pd.isna(c2_content):
                c2_content = "Unknown"
        except Exception as e:
            st.error(f"无法获取C2单元格内容 {uploaded_file.name}: {e}")
            c2_content = "Unknown"

        # 去除非法字符，确保文件名合法
        c2_content = "".join([c for c in str(c2_content) if c.isalnum() or c in (' ', '.', '_')]).rstrip()

        # 如果C2内容为空，则使用默认值
        if not c2_content or c2_content.strip() == "":
            c2_content = f"Unknown_{uploaded_file.name}"

        # 确保文件名唯一
        counter = 1
        base_file_name = c2_content
        while os.path.exists(os.path.join(new_folder_name, f"{c2_content}.xlsx")):
            c2_content = f"{base_file_name}_{counter}"
            counter += 1

        # 重命名文件
        new_file_name = f"{c2_content}.xlsx"
        file_path = os.path.join(new_folder_name, new_file_name)

        # 保存文件到新文件夹
        try:
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
        except Exception as e:
            st.error(f"无法保存文件 {uploaded_file.name}: {e}")
            continue  # 跳过无法保存的文件

    # 压缩新文件夹
    zip_file_name = f"{new_folder_name}.zip"
    shutil.make_archive(new_folder_name, 'zip', new_folder_name)

    # 提供下载链接
    with open(zip_file_name, "rb") as f:
        bytes_data = f.read()
        st.download_button(
            label="下载压缩包",
            data=bytes_data,
            file_name=zip_file_name,
            mime="application/zip"
        )

    # 清理临时文件夹和压缩包
    shutil.rmtree(new_folder_name)
    os.remove(zip_file_name)