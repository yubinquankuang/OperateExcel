import xlwt
from ReadXls import read_excel
from data import *

def write_excel(filename, errname, data):
    myWorkbook = xlwt.Workbook()
    myErrbook = xlwt.Workbook()
    work_row = 1
    err_row = 1
    myErrSheet = myErrbook.add_sheet("缺少订单编号")
    mySheet = myWorkbook.add_sheet("合同信息")
    # myStyle = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')  # 数据格式
    row = len(data)
    col = len(data[0])
    # 表头录入
    for j in range(col):
        myErrSheet.write(0, j, data[0][j])
        mySheet.write(0, j, data[0][j])
    for i in range(1,row):
        if data[i][0]:
            for j in range(col):
                mySheet.write(work_row,j,data[i][j])
            work_row += 1
        else:
            if "合" in data[i][6] or "总" in data[i][6]:
                continue
            for j in range(col):
                myErrSheet.write(err_row, j, data[i][j])
            err_row += 1
    myWorkbook.save(filename)
    myErrbook.save(errname)

def write_final_xls(filename,data):
    myWorkbook = xlwt.Workbook()
    # myStyle = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')  # 数据格式
    for col in COLUMS:
        write_sheet(myWorkbook,data,col['cols'],col['name'],col['func'])
    myWorkbook.save(filename)

def write_sheet(myWorkbook,data,col_list,sheet_name,handleFunc = None):
    """
    写入新的sheet
    myWorkbook : 制定的表格
    data : 要处理的数据
    col_list : 对应的列下标列表
    sheet_name : sheet名称
    """
    if handleFunc:
        data = handleFunc(data)
    mySheet = myWorkbook.add_sheet(sheet_name)
    # myStyle = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')  # 数据格式
    row = len(data)
    for i in range(row):
        for j in range(len(col_list)):
            mySheet.write(i, j, data[i][col_list[j]])





if __name__ == '__main__':
    pass
    # print(write_path,read_path)
    # data = read_excel(read_path,1,4153)
    # write_excel(write_path,err_path,data)
    # data = read_excel(write_path,0,2000)
    # print(data[0])
    # write_final_xls(final_path,data)