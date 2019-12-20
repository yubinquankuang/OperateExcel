import xlwt
import re
import os
from base.baseOperate import *

readF = os.path.join(os.getcwd(), 'read')
writeF = os.path.join(os.getcwd(), 'write')
errF = os.path.join(os.getcwd(), 'err')

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

def matchSheet(numlist, dataSheet, col,outsheet, errsheet):
    '''
    匹配数据，并将匹配好的和没有匹配的放入对应的表中
    numlist： 合同列表（参照）
    dataSheet： 对应表的数据sheet
    col: 匹配对应的列
    outsheet: 对应输出
    errsheet： 筛选过后
    '''
    datas = read_sheet(dataSheet,0)
    err_rows = 0
    out_rows = 0
    write_rows(outsheet, datas[0],out_rows)
    write_rows(errsheet,datas[0],err_rows)
    for data in datas:
        temp = re.findall(r'HT\d{1,}', str(data[col]))
        result = False
        if temp:
            for num in numlist:
                if num in temp:
                    result = True
                    break
        else:
            continue
        if result:
            out_rows += 1
            write_rows(outsheet, data, out_rows)
        else:
            err_rows += 1
            write_rows(errsheet, data, err_rows)

def matchSheet2(numlist, dataSheet, col, outsheet, addData):
    '''
    匹配数据，并将匹配好的和没有匹配的放入对应的表中
    numlist： 合同列表（参照）
    dataSheet： 对应表的数据sheet
    col: 匹配对应的列
    outsheet: 对应输出
    errsheet： 筛选过后
    addData: 要组合的表的数据
    '''
    # 添加表头
    datas = read_sheet(dataSheet,0)
    out_rows = 0
    # temp = []
    # temp.extend(addData[0])
    # temp.extend(datas[0])
    # write_rows(outsheet, temp,out_rows)

    # 匹配写入数据
    for data in datas:
        temp = re.findall(r'HT\d{1,}', str(data[col]))
        current_row = 0
        for num in numlist:
            if num[0] in temp:
                current_row = int(num[1])
                break
        temp = []
        temp.extend(addData[current_row])
        temp.extend(data)
        write_rows(outsheet, temp, out_rows)
        out_rows += 1


def getMatchDatas(numberFile, dataFile, errFile, outFile):
    '''
    numberFile: 获取对应编号列表的文件
    dataFile: 获取数据文件
    errFile: 获取剩余文件
    outFile: 获取筛选文件
    '''
    pass

def getNumList(dataSheet, col):
    datas = read_sheet(dataSheet, 0)
    dataList = []
    for data in datas:
        temp = re.findall(r'HT\d{1,}', str(data[col]))
        dataList.extend(temp)
    return dataList


if __name__ == "__main__":
    filename = os.path.join(readF, '3.xlsx')
    outname = os.path.join(writeF, 'Number.xlsx')
    data = getComNums(filename, outname,2)
    numlist = [x[0] for x in data]
    # 匹配结果并输出
    # 获取表名和sheet
    filename = os.path.join(readF, '1.xlsx')
    dataWork = xlrd.open_workbook(filename)
    names = dataWork.sheet_names()
    print(names)
    # 生成对应的book 和 sheet
    errbook = xlwt.Workbook()
    workbook = xlwt.Workbook()
    cols = [0, 3, 2, 3, 3, 3, 0]
    index = 0

    for name in names:
        dataSheet = dataWork.sheet_by_index(index)
        errsheet = errbook.add_sheet(name)
        worksheet = workbook.add_sheet(name)
        matchSheet(numlist, dataSheet, cols[index], worksheet, errsheet)
        index += 1

    errfile = os.path.join(writeF,"err.xlsx")
    workfile = os.path.join(writeF, 'work.xlsx')
    errbook.save(errfile)
    workbook.save(workfile)
    # +++++++++++++++++++++++

    # 获取新的对应数据
    newCheckfile = os.path.join(readF,"2.xlsx")
    newCheckout = os.path.join(writeF,"newNum.xlsx")
    newCheckData = getComNums(newCheckfile, newCheckout)
    numlist = [x[0] for x in newCheckData[1:]]

    # 重新校验生成新的nework 和 newerr
    filename = os.path.join(writeF, 'err.xlsx')
    dataWork = xlrd.open_workbook(filename)
    names = dataWork.sheet_names()
    # 生成对应的book 和 sheet
    errbook = xlwt.Workbook()
    workbook = xlwt.Workbook()
    cols = [0, 3, 2, 3, 3, 3, 0]
    index = 0

    for name in names:
        dataSheet = dataWork.sheet_by_index(index)
        errsheet = errbook.add_sheet(name)
        worksheet = workbook.add_sheet(name)
        matchSheet(numlist, dataSheet, cols[index], worksheet, errsheet)
        index += 1

    errfile = os.path.join(writeF, "newerr.xlsx")
    workfile = os.path.join(writeF, 'newwork.xlsx')
    errbook.save(errfile)
    workbook.save(workfile)

    # +++++++++++++++++++
    # 表格合并
    addDataFile = os.path.join(readF, '2.xlsx')
    addData = read_excel(addDataFile,0)
    filename = os.path.join(writeF, "newwork.xlsx")
    dataWork = xlrd.open_workbook(filename)
    checkListFile = os.path.join(writeF, "newNum.xlsx")
    checklist = read_excel(checkListFile)

    # 生成对应的数据表格
    workbook = xlwt.Workbook()
    cols = [0, 3, 2, 3, 3, 3, 0]
    index = 0
    for name in names:
        dataSheet = dataWork.sheet_by_index(index)
        worksheet = workbook.add_sheet(name)
        matchSheet2(checklist, dataSheet, cols[index], worksheet, addData)
        index += 1

    workfile = os.path.join(writeF, '合并.xlsx')
    workbook.save(workfile)

    # 生成2N4的筛选列表
    dataFile = workfile
    deleteList = []
    dataWork = xlrd.open_workbook(dataFile)
    index = 0
    for name in names:
        dataSheet = dataWork.sheet_by_index(index)
        deleteList.extend(getNumList(dataSheet,cols[index]))
        index += 1

    # 对2表进行处理
    datafile = os.path.join(readF, "2.xlsx")
    datas= read_excel(datafile,0)
    outfile = os.path.join(writeF,"2-2N4.xlsx")

    workbook = xlwt.Workbook()
    sheet1 = workbook.add_sheet("筛选数据")
    row = 0

    for data in datas:
        temp = re.findall(r'HT\d{1,}', str(data[0]))
        result = False
        for num in deleteList:
            if num in temp:
                result = True
                break
        if result:
            break
        write_rows(sheet1,data,row)
        row += 1
    workbook.save(outfile)

    # ++++++++++++
    # 重新校验生成新的nework 和 newerr
    filename = os.path.join(writeF, 'err.xlsx')
    dataWork = xlrd.open_workbook(filename)
    names = dataWork.sheet_names()
    # 生成对应的book 和 sheet
    errbook = xlwt.Workbook()
    workbook = xlwt.Workbook()
    cols = [0, 3, 2, 3, 3, 3, 0]
    index = 0

    for name in names:
        dataSheet = dataWork.sheet_by_index(index)
        errsheet = errbook.add_sheet(name)
        worksheet = workbook.add_sheet(name)
        matchSheet(deleteList, dataSheet, cols[index], worksheet, errsheet)
        index += 1

    errfile = os.path.join(writeF, "4-2N4.xlsx")
    workfile = os.path.join(writeF, '4N2N4.xlsx')
    errbook.save(errfile)
    workbook.save(workfile)





