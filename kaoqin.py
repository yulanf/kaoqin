import win32com.client as win32

# doctor 6
# master 27

columns = ['姓名','应到','实到','迟到','早退','旷工','加班','工作时间','未签到','未签退','出勤时间']

def trans_to_xlsx():
	fname = "C:\\Users\\wxy\\Desktop\\test\\1.14-1.20.xls.xls"
	excel = win32.gencache.EnsureDispatch('Excel.Application')
	wb = excel.Workbooks.Open(fname)
	
	wb.SaveAs(fname[:-5]+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
	wb.Close()                               #FileFormat = 56 is for .xls extension
	excel.Application.Quit()



import pandas as pd
df = pd.read_excel('1.14-1.20.xlsx')




