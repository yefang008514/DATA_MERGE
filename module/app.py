import streamlit as st
import os
import time
import warnings
import xlwings as xw
import pandas as pd

from module.read_data import read_config_map,read_folder_data_merge_muti


#替换路径名称最后一段
def replace_last_segment(file_path, new_segment):
    # 使用 os.path.split 将路径分割为目录和文件名
    directory, filename = os.path.split(file_path)
    
    # 将文件名替换为新的字符串
    new_file_path = os.path.join(directory, new_segment)
    
    return new_file_path

# 主合并函数
def main_with_streamlit(folder_path, config_path):
    # 1.读取配置文件
    config_dict = read_config_map(config_path)

    # 2.获取各文件夹名称路径
    folder_name_list = os.listdir(folder_path)

    folder_name_list=[n for n in folder_name_list if '.' not in n]


    # 3.遍历文件夹，读取数据并合并
    # 存储日志
    re = []
    for i, x in enumerate(folder_name_list):

        folder_path_run = os.path.join(folder_path, x)
        config_dict_run = config_dict.get(x, {})

        # 显示进度条
        st.write(f"正在处理文件夹: {x} ({i + 1}/{len(folder_name_list)})")
        st.progress((i+1) / len(folder_name_list))
        
        # 读取数据并合并
        try:
            df_x = read_folder_data_merge_muti(folder_path=folder_path_run, config_dict=config_dict_run, engine='pandas')
            # df_x = read_folder_data_merge_muti(folder_path=folder_path_run, config_dict=config_dict_run, engine='xlwings')
            # 保存数据
            save_path = replace_last_segment(file_path=folder_path,new_segment='整理后网银流水_auto')
            os.makedirs(save_path, exist_ok=True)
            save_file_path = os.path.join(save_path, f'{x}.xlsx')
            df_x.to_excel(save_file_path, index=False)

            # 记录日志
            temp_dict = {
                '原始文件夹路径': folder_path_run,
                '保存路径': save_file_path,
                '合并后行数': len(df_x),
                '合并后收入金额': df_x['收入'].sum(),
                '合并后支出金额': df_x['支出'].sum()
            }
            re.append(temp_dict)
        except Exception as e:
            st.error(f"处理文件夹 {x} 时出错: {str(e)}")

    #4.生成日志
    log_df = pd.DataFrame(re)
    log_file_path = os.path.join(save_path, '合并日志.xlsx')
    log_df.to_excel(log_file_path, index=False)

    st.success("合并完成! 日志保存在: " + log_file_path)


def main_ui():
    st.title("文件合并工具")

    # 用户选择文件夹路径
    folder_path = st.text_input("选择包含待处理文件的文件夹路径:")

    # 用户上传配置文件
    config_file = st.file_uploader("上传配置映射表:", type=['xlsx','xlsm'])

    st.markdown(
    '''
    copyright
    © [20250110] [立信会计师事务所浙江分所 21部]。保留所有权利。

    使用本工具遇到任何问题，请联系：[yefang@bdo.com.cn]
    ''')

    # 开始处理按钮
    if st.button("开始处理"):
        if folder_path and config_file:
            with open("temp_config.xlsx", "wb") as f:
                f.write(config_file.getbuffer())

            try:
                main_with_streamlit(folder_path, "temp_config.xlsx")
            finally:
                os.remove("temp_config.xlsx")
        else:
            st.error("请先输入文件夹路径并上传配置文件!")


# Streamlit UI界面
if __name__ == "__main__":

    main_ui()
