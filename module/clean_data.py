import xlwings as xw
import pandas as pd


#处理定位表头，去掉干扰数据
def process_df(data):
    df=data.copy()

    #####寻找表头 表头不能是最后一行#####
    missing_values_count = df.iloc[:-1,:].isnull().sum(axis=1)

    # 找到空值最少的行（表头）
    header_row_index = missing_values_count.idxmin()

    # 提取表头行
    header_row = df.iloc[header_row_index]

    #去掉表头字符串中的空格
    header_row = header_row.apply(lambda x:x.strip() if isinstance(x,str) else x)

    # 重新整理 DataFrame，将表头设置为空值最少的行，并删除该行
    df_body = df.iloc[header_row_index+1:,:]
    # 将表头行设置为 DataFrame 的列名
    df_body.columns = header_row
    # 重新设置索引
    df_body.reset_index(drop=True, inplace=True)

    return df_body


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
    final_dict=config_dict
    raw_col=['无']+list(df.columns)

    #初始化
    amount_col=final_dict['金额列']
    flag_col=final_dict['标识列']
    income_col=final_dict['收入标识']
    expense_col=final_dict['支出标识']

    acct_col=final_dict['户名']
    in_acct_col=final_dict['收款人户名']
    ex_acct_col=final_dict['付款人户名']

    #判断final_dict里面的元素是否都在原始的data里面
    for k,v in final_dict.items():
        if pd.isna(v):
            pass
        elif k in ['收入标识','支出标识']:
            if flag_col=='无':#如果标识列为空，则不用判断
                pass
            elif flag_col not in raw_col:#如果标识列不在原始数据中，则报错
                raise ValueError(f'列[{flag_col}] 不在原始数据中，请检查配置映射表!')
            else:#检查标识列是否有收入和支出标识
                if income_col not in df[flag_col].unique().tolist():
                    raise ValueError(f'标识列[{income_col}] 值不正确，请检查配置映射表!')
                if expense_col not in df[flag_col].unique().tolist():
                    raise ValueError(f'标识列[{expense_col}] 值不正确，请检查配置映射表!')
        else:
            if v not in raw_col:
                raise ValueError(f'列[{v}] 不在原始数据中，请检查配置映射表!')
    

    #1.通过时间频率过滤数据
    df=filter_df_by_date(data=df,col_name=final_dict['时间']).copy()
    

    #2.判断是否有【金额列】
    flag_amount_col=pd.isna(amount_col) #判断金额列是否为空


    #3.处理收入支出通过标识区分的情况
    if (flag_amount_col==False) and (flag_col=='无'): #无标识按正负区分
        df[amount_col]=df[amount_col].apply(lambda x:float(str(x).replace(',','')))
        df['收入']=df[amount_col].apply(lambda x:x if x>0 else 0)
        df['支出']=df[amount_col].apply(lambda x:-x if x<0 else 0)
        final_dict['收入']='收入'
        final_dict['支出']='支出'
    elif (flag_amount_col==False) and (flag_col!='无'):#有标识按标识区分
        #如果金额列里面有逗号需要替换成空格
        if ',' in df[amount_col].astype(str).str.cat(sep=''):
            df[amount_col]=df[amount_col].apply(lambda x:float(str(x).replace(',',''))).copy()
        else:
            pass

        # df['收入']=df.apply(lambda x:abs(x[amount_col]) if x[flag_col]==income_col else 0,axis=1)
        # df['支出']=df.apply(lambda x:abs(x[amount_col]) if x[flag_col]==expense_col else 0,axis=1)

        #替换写法 使用辅助列
        df['income_col']=(df[flag_col]==income_col).astype(int)
        df['expense_col']=(df[flag_col]==expense_col).astype(int)
        df['收入']=df[amount_col]*df['income_col'].abs()
        df['支出']=df[amount_col]*df['expense_col'].abs()
        final_dict['收入']='收入'    
        final_dict['支出']='支出'
    else:
        pass 

    #4.处理【对方户名】需要通过标识区分的情况，如果是收款，对方户名就是付款人
    if acct_col=='无' or pd.isna(final_dict['户名']):
        if flag_col=='无':
            df['户名']=df.apply(lambda x:x[ex_acct_col] if x[amount_col]>0 else x[in_acct_col],axis=1)
            final_dict['户名']='户名'    
        elif flag_col!='无':
            df['户名']=df.apply(lambda x:x[ex_acct_col] if x[flag_col]==income_col else x[in_acct_col],axis=1)
            final_dict['户名']='户名'
        else:
            pass


    #5.仅保留需要的列
    must_col=['时间','收入','支出','余额','户名','摘要']
    col_name_list=[v for k,v in final_dict.items() if k in must_col]
    df=df[col_name_list].copy()
    df.columns=must_col

    #金额保留两位小数
    df['收入']=df['收入'].apply(lambda x:x.strip() if isinstance(x,str) else x)
    df['支出']=df['支出'].apply(lambda x:x.strip() if isinstance(x,str) else x)
    df['余额']=df['余额'].apply(lambda x:x.strip() if isinstance(x,str) else x)

    df['收入']=df['收入'].replace('',None)
    df['支出']=df['支出'].replace('',None)
    df['余额']=df['余额'].replace('',None)

    df['收入']=df['收入'].replace('-',None)
    df['支出']=df['支出'].replace('-',None)
    df['余额']=df['余额'].replace('-',None)

    df['收入'].fillna(0,inplace=True)
    df['支出'].fillna(0,inplace=True)
    df['余额'].fillna(0,inplace=True)

    df['收入']=df['收入'].apply(lambda x:round(float(str(x).replace(',','')),2))
    df['支出']=df['支出'].apply(lambda x:round(float(str(x).replace(',','')),2))
    df['余额']=df['余额'].apply(lambda x:round(float(str(x).replace(',','')),2))

    return df