
import xlwt,xlrd
import os
from datetime import datetime
from xlrd import xldate_as_tuple 
import calendar
path = input() #文件夹目录
files= os.listdir(path) #得到文件夹下的所有文件名称

wb_all = xlwt.Workbook(encoding = 'utf-8')
sheet2 = wb_all.add_sheet('系统负荷曲线') 
sheet2.write(0,0,'系统负荷曲线')
sheet2.write(1,0,'日期')
sheet2.write(1,2,'星期')

times = ['0:00','1:00','2:00','3:00','4:00','5:00','6:00','7:00','8:00','9:00','10:00','11:00','12:00']
times += ['13:00','14:00','15:00','16:00','17:00','18:00','19:00','20:00','21:00','22:00','23:00']
col = 3
for time in times:
    sheet2.write(1,col,time)
    col = col +1

weekStr = ['星期一','星期二','星期三','星期四','星期五','星期六','星期日']
row = 1
col = 3

for file in files: #遍历文件夹
    if os.path.isdir(file): #判断是否是文件夹，不是文件夹才打开
        continue
    if not file.endswith('.xlsx'):
        continue
    if file == 'report.xlsx':
        continue
    if file.startswith('~$'):
        continue
    excel_path = path+"/"+file #打开文件

    print(excel_path)


    wb_temp=xlrd.open_workbook(excel_path) #打开待复制的表
    tmpSheet = wb_temp.sheet_by_index(0)#第一个sheet
    rowCount = tmpSheet.nrows

    curDateStr = ''
    for i in range(1, rowCount):
        dvalue = tmpSheet.cell(i,0).value
        dtuple = xldate_as_tuple(dvalue,0)
        date = datetime(*dtuple)
        dstr = date.strftime('%Y-%m-%d')
        week = calendar.weekday(date.year,date.month,date.day)

        data = tmpSheet.cell(i,1).value

        print(dstr)
        print(data)

        if curDateStr != dstr:
            curDateStr = dstr
            col = 3
            row = row+1
            sheet2.write(row,0,dstr)
            sheet2.write(row,1,weekStr[week])
        sheet2.write(row,col,data)
        col = col+1
    
wb_all.save('report.xls')  # 保存工作簿
