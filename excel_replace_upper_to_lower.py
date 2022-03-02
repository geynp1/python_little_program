import os # 导入os模块
import xlwings as xw # 导入xlwings模块
import re
file_path = 'c:\\1' # 给出要新增工作表的工作簿所在的文件夹路径
file_list = os.listdir(file_path) # 列出文件夹下所有文件和子文件夹的名称
app = xw.App(visible=False,add_book=False) #启动Excel程序
for i in file_list:
    if i.startswith('~$'): # 判断是否有文件名以"~$"
        continue # 如果有，则跳过这种类型的文件
    # print(i)
    sheet_name = re.match(r"(.*?).xlsx",i)# 获取文件名
    real_sheet_name = sheet_name.group(1).lower()# 文件名中的字母小写
    # print(real_sheet_name)
    file_paths = os.path.join(file_path,i)#构造需要新增工作表的工作簿的文件路径
    # print(file_paths)
    workbook = app.books.open(file_paths) # 根据路径打开需要新增工作表的工作簿
    worksheets = workbook.sheets  # 获取工作簿中的所有工作表
    for i in range(len(worksheets)):  # 遍历获取到的工作表
        worksheets[i].name = real_sheet_name  # 重命名工作表
    for j in worksheets:
        value = j['A1'].expand('table').value
        # print(value)
        for index,vals in enumerate(value): # 遍历所有行
            for i in range(len(vals)): # 修改行里的数据
                # print(vals[i])
                if type(vals[i])==str: # 如果类型是字符串
                    # print(type(vals[i]))
                    if str(vals[i]).isalpha(): # 如果是字母
                        vals[i] = vals[i].lower() # 小写替换
                        value[index] = vals # 写入value
            print(vals)
        j['A1'].expand('table').value =value # 写入整个表
    workbook.save('c:\\2\\'+real_sheet_name+".xlsx")  # 另存重命名工作表后的工作簿
app.quit()  # 退出Excel程序