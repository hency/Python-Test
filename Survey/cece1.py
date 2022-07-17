
book = load_workbook(tag_file)   #能写入已存在表中
wb = load_workbook('原始数据.xlsx')
for sheet in wb.sheetnames:
    print(sheet)
    wbsheet=wb[sheet]
    for num in range(3):
        name=wbsheet.cell(1,num*15+10).value
        wbsheet_new = book.create_sheet(name,0)
        wm=list(wbsheet.merged_cells) #开始处理合并单元格形式为“(<CellRange A1：A4>,)，替换掉(<CellRange 和 >,)' 找到合并单元格
        #print (list(wm))
        if len(wm)>0 :
            for i in range(0,len(wm)):
                cell2=str(wm[i]).replace('(<CellRange ','').replace('>,)','')
                #print("MergeCell : %s" % cell2)
                wbsheet_new.merge_cells(cell2)
        for rows in range(40):
            wbsheet_new.row_dimensions[rows+1].height = wbsheet.row_dimensions[rows+1].height
            for col in range(14):
                wbsheet_new.column_dimensions[get_column_letter(col+1)].width = wbsheet.column_dimensions[get_column_letter(col+1)].width
                wbsheet_new.cell(row=rows+1,column=col+1,value=wbsheet.cell(rows+1,num*15+col+1).value)
                if wbsheet.cell(rows+1,num*15+col+1).has_style: #拷贝格式
                    wbsheet_new.cell(row=rows+1,column=col+1).font = copy(wbsheet.cell(rows+1,num*15+col+1).font)
                    wbsheet_new.cell(row=rows+1,column=col+1).border = copy(wbsheet.cell(rows+1,num*15+col+1).border)
                    wbsheet_new.cell(row=rows+1,column=col+1).fill = copy(wbsheet.cell(rows+1,num*15+col+1).fill)
                    wbsheet_new.cell(row=rows+1,column=col+1).number_format = copy(wbsheet.cell(rows+1,num*15+col+1).number_format)
                    wbsheet_new.cell(row=rows+1,column=col+1).protection = copy(wbsheet.cell(rows+1,num*15+col+1).protection)
                    wbsheet_new.cell(row=rows+1,column=col+1).alignment = copy(wbsheet.cell(rows+1,num*15+col+1).alignment)
wb.close()
book.save('拆分后表.xlsx')
book.close()
import xlwings as xw
import os
file_path = 'G:\\KPI考核\\2021\\'
file_list = os.listdir(file_path)
app = xw.App(visible = True,add_book = False) #启动Exel程序，但是不新建工作簿
for i in file_list:
    if i.startswith('~$'):
        continue
    file_paths = os.path.join(file_path, i) #构造需要打印工作表的工作簿的文件路径
    workbook = app.books.open(file_paths) #根据路径打开需要打印工作表的工作簿
    workbook.api.PrintOut() #打印要打印的工作簿
app.quit()