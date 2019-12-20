import xlrd
import xlwt
import re

def read_excel(filename,start = 1,end = 0, sheetIndex = 0, cols = []):
    """
    读取指定excel的行列
    filename: 文件名
    start: 起始位置，行
    end: 结束位置，行
    sheetIndex: 表格下标
    """
    workbook = xlrd.open_workbook(filename)

    sheet_name = workbook.sheet_names()[sheetIndex]
    print("sheet_name: ",sheet_name)
    # sheet = workbook.sheet_by_name(sheet_name)
    sheet = workbook.sheet_by_index(sheetIndex)
    # 统计行列
    sheet_row = sheet.nrows
    sheet_col = sheet.ncols

    # 表格数据读取
    data = []
    if end > sheet_row or end == 0:
        end = sheet_row
    if cols:
        for i in range(start, end):
            temp = []
            for col in cols:
                temp.append(sheet.cell_value(i,col))
            data.append(temp)
    else:
        for i in range(start, end):
            data.append(sheet.row_values(i, 0, sheet_col))
    return data

def read_sheet(sheet,start = 1,end = 0,cols = []):
    """
    读取指定excel的行列
    sheet: sheet对象
    start: 起始位置，行
    end: 结束位置，行
    sheetIndex: 表格下标
    """
    # 统计行列
    sheet_row = sheet.nrows
    sheet_col = sheet.ncols

    # 表格数据读取
    data = []
    if end > sheet_row or end == 0:
        end = sheet_row
    if cols:
        for i in range(start, end):
            temp = []
            for col in cols:
                temp.append(sheet.cell_value(i,col))
            data.append(temp)
    else:
        for i in range(start, end):
            data.append(sheet.row_values(i, 0, sheet_col))
    return data

def write_rows(sheet, datas, row, cols = []):
    """
    sheet: 指定的sheet
    datas: 要写入的行数据
    """
    ncol = len(datas)
    ncols = len(cols)
    # print("shuju",ncol, ncols)
    if cols:
        for i in range(ncol):
            if cols[i] == -1:
                continue
            try:
                sheet.write(row, cols[i], datas[i])
            except:
                print(i,cols[i],len(datas),ncols,ncol)

    else:
        for i in range(ncol):

            sheet.write(row, i, datas[i])

def write_cols(sheet, datas, col):
    """
    sheet: 指定的sheet
    datas: 要写入的列数据
    """
    nrow = len(datas)
    for i in range(nrow):
        sheet.write(i, col, datas[i])


def write_sheet(sheet, datas, cols = [],start = 0):
    """
    将datas写入指定的sheet中，指定从表格的第几行写入
    sheet: 指定的sheet
    data: 要写入的表数据，
    cols: 写入对应的列的下标
    start: 开始写入的行的位置，base=0
    """
    nrow = len(datas)
    index = 0
    for i in range(start,nrow+start):
        write_rows(sheet, datas[index], i, cols)
        index += 1

def write_file(file, sheet, datas, cols = [],start = 0):
    """
    生成文件，excel
    file: 文件名
    sheet: 指定的sheet
    data: 要写入的表数据，
    cols: 写入对应的列的下标
    start: 开始写入的行的位置，base=0
    """
    work = xlwt.Workbook()
    sheet1 = work.add_sheet(sheet)
    write_sheet(sheet1, datas, cols, start)
    work.save(file)

def getComNums(filename,outname,chooseCol=0):
    '''
    获取编号列表
    filename: 文件名

    return : 合同列表
    '''
    data = read_excel(filename)
    new_data = []
    col = 1
    new_data.append(["合同号", '对应行'])
    for i in data:
        temp = re.findall(r'HT\d{1,}',str(i[chooseCol]))
        if temp:
            for t in temp:
                new_data.append([t, col])
        col += 1
    book = xlwt.Workbook()
    sheet = book.add_sheet('合同编号表')

    write_sheet(sheet, new_data)
    book.save(outname)
    return new_data
