# doctor 6
# master 27

columns = ['部门','姓名','应到','实到','迟到','早退','旷工','加班','工作时间','未签到','未签退','出勤时间', '加权出勤', 'ratio']

def trans_to_xlsx():
  import win32com.client as win32
  fname = "C:\\Users\\wxy\\Desktop\\test\\1.14-1.20.xls.xls"
  excel = win32.gencache.EnsureDispatch('Excel.Application')
  wb = excel.Workbooks.Open(fname)
  
  wb.SaveAs(fname[:-5]+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
  wb.Close()                               #FileFormat = 56 is for .xls extension
  excel.Application.Quit()



import pandas as pd
import numpy as np

def select_req_col():
    '''选取报表列'''
    dd = pd.read_excel('1.14-1.20.xlsx')
    df = pd.DataFrame()
    
    for item in columns[:-2]:
        df[item] = dd[item]
    
    doctor_jq = df['出勤时间']+df['加班']-df['旷工']*7-df['早退']-df['迟到']
    master_jq = df['出勤时间']+df['加班']-df['旷工']*6.5-df['早退']-df['迟到']
    df['加权出勤'] = np.where(df['部门']=='博士', doctor_jq, master_jq)
    
    doctor_ratio = df['工作时间']/(df['应到']*7)
    master_ratio = df['工作时间']/(df['应到']*6.5)
    df['ratio'] = np.where(df['部门']=='博士', doctor_ratio, master_ratio)

    writer = pd.ExcelWriter('1.14-1.20.xlsx')
    dd.to_excel(writer, 'origin', index=None)
    df.to_excel(writer, 'ver1', index=None)
    writer.save()


def split_by_department():
    # 读取文件准备按部门分离
    df = pd.read_excel('1.14-1.20.xlsx', 'ver1')
    sid = df['部门']
    index1 = 0
    index2 = 0
    flag = True
    for index, i in enumerate(sid):
        # print(index,i)
        if i == '研二' and flag == True:
            # print(ind ex)
            index1 = index
            flag = False
        if i == '研一':
            index2 = index
            break
    
    # 按部门拆分df
    # dd = pd.DataFrame()
    df_doctor = df.loc[:index1-1]
    # dt = pd.DataFrame()
    df_master2 = df.loc[index1:index2-1]
    # dz = pd.DataFrame()
    df_master1 = df.loc[index2:]
    
    # 分别排序
    df_doctor = df_doctor.sort_values(by=['ratio','加权出勤'] ,ascending=(False, False))
    df_master2 = df_master2.sort_values(by=['ratio','加权出勤'], ascending=(False, False))
    df_master1 = df_master1.sort_values(by=['ratio','加权出勤'], ascending=(False, False))

    # 合并为一个
    df = pd.concat([df_doctor,df_master2,df_master1])


    df.loc[index1-1]=None
    df.loc[index1] = columns

    df.loc[index2+1]=None
    df.loc[index2+2]=columns

    writer = pd.ExcelWriter('1.14-1.20_split.xlsx')
    dd.to_excel(writer, '博士', index=None)
    dt.to_excel(writer, '研二', index=None)
    dz.to_excel(writer, '研一', index=None)
    writer.save()


def a():
    # 读取3个sheet
    df_doctor = pd.read_excel('1.14-1.20_split.xlsx', '博士')
    df_master2 = pd.read_excel('1.14-1.20_split.xlsx', '研二')
    df_master1 = pd.read_excel('1.14-1.20_split.xlsx', '研一')
    
    
    # 合并为一个
    
    writer = pd.ExcelWriter('1.14-1.20_report.xlsx')
    df.to_excel(writer, 'v1', index=None)
    writer.save()
    
    
    df = pd.read_excel('1.14-1.20_report.xlsx', 'v1')
    sid = df['部门']
    flag = 1
    # 加表头
    for index, i in enumerate(sid):
        # print(index,i)
        if i == '研二' and flag == 1:
            # print(index)
            df.loc[index-1]=columns
            flag = 2
    
        if i == '研一':
            df.loc[index]=columns
            break
    # 加空行
    flag = 1
    for index, i in enumerate(sid):
        # print(index,i)
        if i == '研二' and flag == 1:
            # print(index)
            df.loc[index-2]=None
            flag = 2
    
        if i == '研一':
            df.loc[index-2]=None
            break


    writer = pd.ExcelWriter('1.14-1.20_report_f.xlsx')
    df.to_excel(writer, 'v1', index=None)
    writer.save()           
    

if __name__ == '__main__':
    # trans_to_xlsx()
    select_req_col()
    # add_cal_col()
