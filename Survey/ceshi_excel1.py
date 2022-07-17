
import xlwings as xw
app = xw.App(visible=False, add_book=False)
file_path = "D:\\2022\\基坑监测\\大剧院\\数据库\\cx拆\\测斜仿真逆向运算CX10.xlsx"
workbook = app.books.open(file_path)
worksheet = workbook.sheets
for i in worksheet:  # 遍历工作簿中所有工作表
    new_workbook = app.books.add()  # 新建工作簿
    new_worksheet = new_workbook.sheets[0]  # 选中新建工作簿中的第1张工作表
    i.copy(before=new_worksheet)  # 将原来工作簿中的当前工作表复制到新建工作簿的第1张工作表之前
    new_workbook.save("D:\\2022\\基坑监测\\大剧院\\数据库\\cx拆\\{}.xlsx".format(i.name))  # 保存新工作簿
    new_workbook.close()  # 关闭新建工作簿
app.quit()
