import xlwings as xw
import pandas as pd
from multiprocessing.dummy import Pool as ThreadPool
import os
import warnings

from module.clean_data import clean_df,process_df


########################################读df##########################################

# 读取配置映射表信息
def read_config_map(config_path):
    df = pd.read_excel(config_path,dtype=object,sheet_name='配置映射表')
    #去掉字符串前后的空格
    df=df.applymap(lambda x:x.strip() if isinstance(x,str) else x)
    #把映射表转换成字典
    result = df.set_index('银行').to_dict(orient='index')
    return result

# 根据文件路径读取文件返回df 
def read_file(file_path):
    warnings.filterwarnings("ignore")
    if file_path.endswith('.xlsx'):
        df = pd.read_excel(file_path,engine='openpyxl',dtype=object,header=None)
    elif file_path.endswith('.xls'):
        df = pd.read_excel(file_path,engine='xlrd',dtype=object,header=None)

    elif file_path.endswith('.csv'):
            for encoding in ['ansi', 'gbk', 'utf-8']:
                try:
                    df = pd.read_csv(file_path, dtype=object, header=None, encoding=encoding)
                    break
                except Exception:
                    continue
            else:
                try:
                    #不自动识别表头
                    df = read_data_xlwings(path=file_path,auto_header=1)
                except Exception:
                    raise ValueError(f"无法解析,请检查文件编码格式,尝试保存成xlsx格式: {file_path} ")
    else:
        print(f"该文件格式不支持: {file_path}")
        return None

    df=process_df(data=df).copy()
    return df 

# 使用xlwings读取数据 返回df  auto_header 如果随便填就不识别表头
def read_data_xlwings(path,sheet_name=None,header=None,auto_header=None):
    mypath=path
    header_final=header if header is not None else 0
    with xw.App(visible=False) as app:
        book=app.books.open(mypath)
        if sheet_name is not None:
            table=book.sheets[sheet_name].used_range
        else:
            table=book.sheets[0].used_range
            df=table.options(pd.DataFrame, header=header_final, index=False).value
        book.close()
    #默认自动识别表头
    if auto_header is None:
        df=process_df(data=df).copy()
        return df
    else:
        return df


########################################读文件路径##########################################

# 选择文件夹路径返回文件夹路径 暂时用不到
def select_folder():
    shell = win32com.client.Dispatch("Shell.Application")
    folder = shell.BrowseForFolder(0, "选择文件夹", 0, 0)
    
    if folder:
        folder_path = folder.Items().Item().Path
        return folder_path
    else:
        return None

#穿透遍历文件夹所有xlsx,xls,csv获取他们的路径保存到列表中,用~$判断非打开的文件
def get_file_list(folder_path):
    file_list = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if (file.endswith('.xlsx') or file.endswith('.xls') or file.endswith('.csv')) and '~$' not in file:
                file_list.append(os.path.join(root, file))
            else:
                continue
    return file_list

########################################多线程合并########################################

#读取数据并合并，再通过字段清洗
def read_folder_data_merge_muti(folder_path,config_dict,engine):

    pool = ThreadPool(6)
    file_list = get_file_list(folder_path)
    #读取数据
    if engine=='xlwings':
        results=pool.map(read_data_xlwings,file_list)
    elif engine=='pandas':
        results=pool.map(read_file,file_list)
    else:
        raise ValueError('引擎参数错误')
    #合并数据
    df_all=pd.concat(results)
    #清洗数据
    result=clean_df(data=df_all,config_dict=config_dict).copy()

    return result