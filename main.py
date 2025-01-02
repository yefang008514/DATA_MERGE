import os 
import pandas as pd 
import xlwings as xw
from multiprocessing import Pool, freeze_support
import win32com.client
import warnings
import time
import sys


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

# 选择文件夹路径返回文件夹路径
def select_folder():
    shell = win32com.client.Dispatch("Shell.Application")
    folder = shell.BrowseForFolder(0, "选择文件夹", 0, 0)
    
    if folder:
        folder_path = folder.Items().Item().Path
        return folder_path
    else:
        return None

#穿透遍历文件夹所有xlsx,xls,csv获取他们的路径保存到列表中
def get_file_list(folder_path):
    file_list = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.xlsx') or file.endswith('.xls') or file.endswith('.csv'):
                file_list.append(os.path.join(root, file))
    return file_list

#处理定位表头，去掉干扰数据
def process_df(data):
    df=data.copy()

    #####寻找表头#####
    missing_values_count = df.isnull().sum(axis=1)
    # 找到空值最少的行（表头）
    header_row_index = missing_values_count.idxmin()
    # 提取表头行
    header_row = df.iloc[header_row_index]
    # 重新整理 DataFrame，将表头设置为空值最少的行，并删除该行
    df_body = df.iloc[header_row_index+1:,:]
    # 将表头行设置为 DataFrame 的列名
    df_body.columns = header_row
    # 重新设置索引
    df_body.reset_index(drop=True, inplace=True)

    return df_body

# 使用xlwings读取数据  auto_header 如果随便填就不识别表头
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

#通过时间频率筛选数据
def filter_df_by_date(data,col_name):
    df_final=data.copy()

    #统计col_name列的字符串长度频率,筛选出现频率最高的长度
    df_final['date_len']=df_final[col_name].astype(str).apply(len)
    col_name_len_freq = df_final['date_len'].value_counts()
    max_len = col_name_len_freq.idxmax()
    
    df_final=df_final[df_final['date_len']==max_len]
    df_final.drop('date_len',axis=1,inplace=True)
    
    return df_final


#根据字典配置清洗数据
def clean_df(data,config_dict):
    df=data.copy()
    final_dict=config_dict.copy()

    #1.通过时间频率过滤数据
    df=filter_df_by_date(data=df,col_name=final_dict['时间']).copy()

    #2.判断是否有【金额列】
    amount_col=final_dict['金额列']
    flag_col=final_dict['标识列']
    income_col=final_dict['收入标识']
    expense_col=final_dict['支出标识']
    flag_amount_col=pd.isna(amount_col) #判断金额列是否为空

    #3.处理收入支出通过标识区分的情况
    if (flag_amount_col==False) and (flag_col=='无'): #无标识按正负区分
        df[amount_col]=df[amount_col].apply(lambda x:float(str(x).replace(',','')))
        df['收入']=df[amount_col].apply(lambda x:x if x>0 else 0)
        df['支出']=df[amount_col].apply(lambda x:-x if x<0 else 0)
        final_dict['收入']='收入'
        final_dict['支出']='支出'
    elif (flag_amount_col==False) and (flag_col!='无'):#有标识按标识区分
        df['收入']=df.apply(lambda x:abs(x[amount_col]) if x[flag_col]==income_col else 0,axis=1)
        df['支出']=df.apply(lambda x:abs(x[amount_col]) if x[flag_col]==expense_col else 0,axis=1)
        final_dict['收入']='收入'    
        final_dict['支出']='支出'
    else:
        pass 


    #4.仅保留需要的列
    must_col=['时间','收入','支出','余额','户名','摘要']
    col_name_list=[v for k,v in final_dict.items() if k in must_col]
    df=df[col_name_list].copy()
    df.columns=must_col

    #金额保留两位小数
    df['收入']=df['收入'].apply(lambda x:round(float(str(x).replace(',','')),2))
    df['支出']=df['支出'].apply(lambda x:round(float(str(x).replace(',','')),2))
    df['余额']=df['余额'].apply(lambda x:round(float(str(x).replace(',','')),2))

    return df

# 读取配置映射表信息
def read_config_map(config_path):
    df = pd.read_excel(config_path,dtype=object)
    #把映射表转换成字典
    result = df.set_index('银行').to_dict(orient='index')
    return result


#读取数据并合并，再通过字段清洗
def read_folder_data_merge_muti(folder_path,config_dict,engine):

    pool = Pool(6)
    file_list = get_file_list(folder_path)

    if engine=='xlwings':
        results=pool.map(read_data_xlwings,file_list)
    elif engine=='pandas':
        results=pool.map(read_file,file_list)
    else:
        raise ValueError('引擎参数错误')

    df_all=pd.concat(results)
    result=clean_df(data=df_all,config_dict=config_dict).copy()

    return result

#进度条
def progress_bar(current, total, bar_length=50):
    """
    显示一个简单的进度条。
    :param current: 当前进度值
    :param total: 总进度值
    :param bar_length: 进度条的长度（字符数）
    """
    percent = current / total
    hashes = '#' * int(percent * bar_length)
    spaces = ' ' * (bar_length - len(hashes))
    sys.stdout.write(f"\rProgress: [{hashes}{spaces}] {percent:.2%}")
    sys.stdout.flush()


def main(path,config_path):
    #读配置文件
    config_dict=read_config_map(config_path)

    #获取各文件夹名称路径
    folder_name_list = os.listdir(path)

    #存储日志
    re=[]

    # 遍历文件夹，读取数据并合并
    for i,x in enumerate(folder_name_list):

        folder_path_run=os.path.join(path,x)
        config_dict_run=config_dict[x]

        #显示进度条
        progress_bar(i, len(folder_name_list))
        
        #读取数据并合并
        print(f'\n正在读取【{x}】数据')
        df_x=read_folder_data_merge_muti(folder_path=folder_path_run,config_dict=config_dict_run,engine='pandas')
        
        
        #保存数据
        print(f'正在保存【{x}】数据')
        save_path=os.path.join(path.replace('原始网银流水','整理后网银流水_auto'))
        os.makedirs(save_path,exist_ok=True)
        save_file_path=os.path.join(save_path,f'{x}.xlsx')
        df_x.to_excel(save_file_path,index=False)
        #记录日志
        temp_dict={'原始文件夹路径':folder_path_run,
                   '保存路径':save_file_path,
                   '合并后行数':len(df_x),
                   '合并后收入金额':df_x['收入'].sum(),
                   '合并后支出金额':df_x['支出'].sum()}
        re.append(temp_dict)
    
    log_df=pd.DataFrame(re)
 
    #保存日志 带时间戳
    # log_df.to_excel(os.path.join(path.replace('原始网银流水','整理后网银流水_auto'),
    # f'合并日志_{time.strftime("%Y%m%d%H%M%S", time.localtime())}.xlsx'),index=False)

    #保存日志 不带时间戳
    log_df.to_excel(os.path.join(path.replace('原始网银流水','整理后网银流水_auto'),'合并日志.xlsx'),index=False)

    print(f'######合并完成#####详见{path.replace("原始网银流水","整理后网银流水_auto")}文件夹')


if __name__ == '__main__':

    
    freeze_support()
    path=r'D:\audit_project\DATA_merge\原始网银流水'
    config_path=r'D:\audit_project\DATA_merge\配置映射表.xlsx'
    main(path,config_path)


