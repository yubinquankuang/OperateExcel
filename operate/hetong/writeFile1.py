import xlwt
import re
import os
from base.baseOperate import *
from base.dataFilter import getDuplicate
from operate.hetong.paramData import cols, duplicate,readf,writef,cols7

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
    err_rows = -1
    out_rows = -1
    # write_rows(outsheet, datas[0],out_rows)
    # write_rows(errsheet,datas[0],err_rows)
    for data in datas:
        temp = re.findall(r'HT\d{1,}', str(data[col]))
        result = False
        if temp:
            for num in numlist:
                if num in temp:
                    result = True
                    break
        else:
            err_rows += 1
            write_rows(errsheet, data, err_rows)
            continue
        if result:
            out_rows += 1
            write_rows(outsheet, data, out_rows)
        else:
            err_rows += 1
            write_rows(errsheet, data, err_rows)

def matchSheet2(numlist, dataSheet, col, outsheet, addData):
    '''
    合并数据
    numlist： 合同列表（参照）
    dataSheet： 对应表的数据sheet
    col: 匹配对应的列
    outsheet: 对应输出
    addData: 要组合的表的数据
    '''
    datas = read_sheet(dataSheet,0)
    out_rows = 0

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


def getMatchDatas(numlist, dataFile, errFile, outFile, cols, handler):
    '''
    numlist: 作为筛选的订单列表
    dataFile: 获取数据文件
    errFile: 未匹配订单的文件名称
    outFile: 已匹配订单的文件名曾
    cols: 各个sheet要匹配的列
    '''
    # 打开要进行筛选的数据表格
    dataWork = xlrd.open_workbook(dataFile)
    names = dataWork.sheet_names()

    # 生成对应的book 和 sheet
    errbook = xlwt.Workbook()
    workbook = xlwt.Workbook()
    index = 0

    for name in names:
        dataSheet = dataWork.sheet_by_index(index)
        errsheet = errbook.add_sheet(name)
        worksheet = workbook.add_sheet(name)
        handler(numlist, dataSheet, cols[index], worksheet, errsheet)
        index += 1

    errbook.save(errFile)
    workbook.save(outFile)

def getNumberMatch(filename, checklist, outname):
    """
    根据获得的合同号列表进行筛选，获取与checklist不必配的数据集合
    filename: 要处理的文件名，如2.xlsx
    checklist: 匹配列表
    outname: 不匹配结果合集的输出文件名，如：2-2N4.xlsx
    """
    datas = read_excel(filename, 0)
    workbook = xlwt.Workbook()
    sheet1 = workbook.add_sheet("筛选数据")
    row = 0

    for data in datas:
        result = False
        # 此处使用的是字符串中内容的匹配
        for num in checklist:
            if num in str(data[0]):
                result = True
                break
        if result:
            continue
        write_rows(sheet1, data, row)
        row += 1
    workbook.save(outname)


def getNumList(dataSheet, col):
    '''
    获取单张表中对应的合同号列表
    dataSheet: 对应的sheet表
    col: 要处理的列号列表，base=0，如 1
    '''
    datas = read_sheet(dataSheet, 0)
    dataList = []
    for data in datas:
        temp = re.findall(r'HT\d{1,}', str(data[col]))
        dataList.extend(temp)
    return dataList

def getSheetsNum(workFile, cols):
    """
    获取如1.xlsx中的合同号列表,有多张sheet子表
    workFile: 要处理的文件名
    cols: 每张子表中要提取的列号，base=0,如[1,1,2]
    """
    deleteList = []
    dataWork = xlrd.open_workbook(workFile)
    index = 0
    for col in cols:

        dataSheet = dataWork.sheet_by_index(index)
        deleteList.extend(getNumList(dataSheet, col))
        print("delete_",index,dataSheet.name,"::",len(deleteList))
        index += 1
    return deleteList

def getDuplicateSheets(sheetIn, sheetOut, cols):
    '''
    sheet表数据去重
    sheetIn: 等待去重的表
    sheetOut: 去重完成后文件输出的表
    cols:
    '''
    dataIn = read_sheet(sheetIn, 0)
    dataOut = getDuplicate(dataIn, cols)
    write_sheet(sheetOut, dataOut)

def getDuplicateFiles(fileIn, fileOut, cols):
    """
    文件去重
    """
    print(fileOut)
    workIn = xlrd.open_workbook(fileIn)
    workOut = xlwt.Workbook()
    names = workIn.sheet_names()
    index = 0

    for name in names:
        sheetIn = workIn.sheet_by_name(name)
        sheetOut = workOut.add_sheet(name)
        getDuplicateSheets(sheetIn, sheetOut, cols[index])
        index += 1

    workOut.save(fileOut)

if __name__ == "__main__":
    cols = [0, 3, 2, 3, 3, 3, 0]

    # 获取表3的合同编号列表
    numFileIn1 = os.path.join(readF, '3.xlsx')
    numFileOut1 = os.path.join(writeF, 'num3.xlsx')
    data = getComNums(numFileIn1, numFileOut1, 2)
    numlist3 = [x[0] for x in data]
    numlist3 = list(set(numlist3))

    # 匹配结果并输出
    # 获取表名和sheet
    dataFile1 = os.path.join(readF, '1.xlsx')
    # 对1表进行去重操作
    dataFile1d = os.path.join(readF, '1d.xlsx')
    getDuplicateFiles(dataFile1, dataFile1d, duplicate)

    errFile1 = os.path.join(writeF,"1-3.xlsx")
    workFile1 = os.path.join(writeF, '1N3.xlsx')
    getMatchDatas(numlist3, dataFile1, errFile1, workFile1, cols, matchSheet)
    names = xlrd.open_workbook(dataFile1).sheet_names()

    # +++++++++++++++++++++++++
    # 生成1-3的numlist
    deleteList13 = getSheetsNum(errFile1, cols)
    deleteList13 = list(set(deleteList13))
    print("1-3List: ", len(deleteList13))
    deleteFile13 = writef("1-3num.xls")
    deleteBook13 = xlwt.Workbook()
    deleteSheet13 = deleteBook13.add_sheet("list")
    write_cols(deleteSheet13, deleteList13, 0)
    deleteBook13.save(deleteFile13)


    # +++++++++++++++++++++++

    # 获取新的对应数据
    numFileIn2 = os.path.join(readF, "2.xlsx")
    numFileOut2 = os.path.join(writeF, "num2.xlsx")
    newCheckData = getComNums(numFileIn2, numFileOut2)
    numlist2 = [x[0] for x in newCheckData[1:]]
    numlist2 = list(set(numlist2))

    # 重新校验生成新的nework 和 newerr
    dataFile2 = os.path.join(writeF, '1-3.xlsx')
    errFile2 = os.path.join(writeF, "4-2.xlsx")
    workFile2 = os.path.join(writeF, '4N2.xlsx')
    getMatchDatas(numlist2,dataFile2,errFile2,workFile2,cols, matchSheet)

    # +++++++++++++++++++
    # 表格合并
    addData = read_excel(numFileIn2,0)
    dataWork3 = xlrd.open_workbook(workFile2)
    addChecklist = read_excel(numFileOut2)

    # 生成对应的数据表格
    workbook = xlwt.Workbook()
    index = 0
    for name in names:
        dataSheet = dataWork3.sheet_by_index(index)
        worksheet = workbook.add_sheet(name)
        matchSheet2(addChecklist, dataSheet, cols[index], worksheet, addData)
        index += 1

    workFile3 = os.path.join(writeF, '5.xlsx')
    workbook.save(workFile3)

    # 生成2N4的筛选列表
    deleteList = getSheetsNum(workFile3, cols7)
    deleteList = list(set(deleteList))
    deleteFile = writef("2N4num.xls")
    deleteBook = xlwt.Workbook()
    deleteSheet = deleteBook.add_sheet("list")
    write_cols(deleteSheet, deleteList, 0)
    deleteBook.save(deleteFile)

    # ++++++++++++
    # 重新校验生成新的nework 和 newerr
    errFile3 = os.path.join(writeF, "4-2N4.xlsx")
    workFile3 = os.path.join(writeF, '4N2N4.xlsx')
    getMatchDatas(deleteList, dataFile2, errFile3, workFile3, cols, matchSheet)

    # 对2表进行处理
    new_list = []
    new_list.extend(deleteList)
    new_list.extend(numlist3)
    data2 = read_excel(numFileIn2, 0)
    workFile4 = os.path.join(writeF, "2-2N4.xlsx")
    getNumberMatch(numFileIn2, new_list, workFile4)





