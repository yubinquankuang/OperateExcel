import xlrd
filename = r'data/read/2017.xlsx'

def read_excel(filename,start = 1,end = 0):
    workbook = xlrd.open_workbook(filename)

    sheet_name = workbook.sheet_names()[0]

    sheet = workbook.sheet_by_name(sheet_name)

    sheet_row = sheet.nrows
    sheet_col = sheet.ncols

    # 表格数据
    # 读取所有数据进行去重
    data = []
    if end == 0:
        return []
    if end > sheet_row:
        end = sheet_row
    for i in range(start, end):
        data.append(sheet.row_values(i, 0, sheet_col))
    return data