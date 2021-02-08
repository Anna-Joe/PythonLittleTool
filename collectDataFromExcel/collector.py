
import xlwt,xlrd
import os

path = input('请输入文件目录：') #文件夹目录
keyWord = input('请输入统计字段：')
files= os.listdir(path) #得到文件夹下的所有文件名称

wb_all = xlwt.Workbook(encoding = 'utf-8')

for file in files: #遍历文件夹
    if os.path.isdir(file): #判断是否是文件夹，不是文件夹才打开
        continue
    if not file.endswith('.xls'):
        continue
    if file == 'sortedData.xls':
        continue
    excel_path = path+"/"+file #打开文件

    sheet2 = wb_all.add_sheet(file) 
    k = 1

    wb_temp=xlrd.open_workbook(excel_path) #打开待复制的表
    tmpSheets = wb_temp.sheet_names()
    for sheetName in tmpSheets:#遍历每个sheet
        sheet = wb_temp.sheet_by_name(sheetName) 
        
        i = 10
        if sheet.nrows < i:
            continue
        
        j = 27
        data = 0
        flag = False
        while i < sheet.nrows:
            # print(excel_path + sheetName)
            # print(i)
            # print(j)
            content = sheet.cell(i,j).value
            if content == (keyWord):
                data = sheet.cell(i,j+1).value
                print(content)
                print( data)
                flag = True
                break
            i = i+1
        if not flag:
            print('未找到关键字'+keyWord+'，请检查'+file)
            continue
        sheet2.write(k,0,sheetName+keyWord)
        sheet2.write(k,1,data)
        k = k+1
    
wb_all.save('sortedData.xls')  # 保存工作簿
