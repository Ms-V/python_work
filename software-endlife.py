import xlrd
import pandas
import openpyxl
import time
import win32com.client as win32
xlsx_file=pandas.read_excel('D:\\pmail\\source\\1超期.xlsx',engine='openpyxl')#打开文件
data_array_notnull_row=xlsx_file.dropna()#去掉有空值得行
data_array_notnull_row=pandas.DataFrame(data=data_array_notnull_row)#格式从Series转换为DataFrame
data_array_notnull_row.to_excel('D:\\pmail\\out\\out_life_'+time.strftime("%Y%m%d%H%M%S",time.localtime())+'.xlsx')#生成去掉空值行的excel方便比对
outlook=win32.Dispatch('Outlook.Application')
#print(len(data_array_notnull_row))
for i in range(len(data_array_notnull_row)):
    mail_item=outlook.CreateItem(0)
    mail_item.Recipients.Add(str(data_array_notnull_row.iat[i,2]))
    mail_item.Subject='软件生命周期终止'
    mail_item.BodyFormat=2
    mail_item.HTMLBody='<body><font size="4">'+data_array_notnull_row.iat[i,4]+',您好<br/>  您的主机'+data_array_notnull_row.iat[i,0]+'中的\
    软件：'+data_array_notnull_row.iat[i,3]+'生命周期已经终止，请您尽快卸载该软件。\
    <br/>  如果卸载需要管理员权限或对软件生命周期和卸载有异议，请联系当地情管部门。<br/><br/>感谢您的支持<br/>情管部-XXX</font></body>'
    mail_item.Send()
