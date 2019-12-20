import os
import xlwt
import xlrd
import time
from operate.hetong.paramData import col4, col5, col2
from base.baseOperate import *

def combinSheets(fileIn,fileOut,cols):
    '''
    合并表格，将子表合成统一表头
    fileIn: 含有多张sheet的，但是数据格式不同
    fileOut: 含有多张sheet表，对应的数据分配再指定的列上
    cols: 各个sheet表列重新分配的对应标签列
    '''
    # 5 表合并
    workIn = xlrd.open_workbook(fileIn)
    workOut = xlwt.Workbook()
    names = workIn.sheet_names()
    new_data = []
    index = 0

    # 完成列的重新排序
    for name in names:
        sheetIn = workIn.sheet_by_name(name)
        sheetOut = workOut.add_sheet(name)
        data = read_sheet(sheetIn, 0)
        print(cols[index])
        write_sheet(sheetOut, data, cols[index])
        index += 1
    workOut.save(fileOut)

    # 获取所有的sheet表的内容并存入new_data
    work = xlrd.open_workbook(fileOut)
    for name in names:
        sheet = work.sheet_by_name(name)
        # 排除表头
        data = read_sheet(sheet, 1)
        new_data.extend(data)
    return new_data


if __name__ == "__main__":
    file5 = os.path.join(os.getcwd(),'write','5.xlsx')
    file4 = os.path.join(os.getcwd(),'write','4-2N4.xlsx')
    file2 = os.path.join(os.getcwd(), 'write', '2-2N4.xlsx')

    file5Out = os.path.join(os.getcwd(),'write','5Out.xlsx')
    file4Out = os.path.join(os.getcwd(),'write','4-2N4Out.xlsx')
    file2Out = os.path.join(os.getcwd(), 'write', '2-2N4Out.xlsx')

    # 5 表合并
    fiveData = combinSheets(file5, file5Out,col5)
    fiveFile = os.path.join(os.getcwd(),'write','5com.xlsx')

    # 4 表合并
    fourData = combinSheets(file4, file4Out,col4)
    fourFile = os.path.join(os.getcwd(), 'write', '4-2N4com.xlsx')

    # 获取2表数据
    twoData = read_excel(file2, 0)

    # 将不同表格的数据写入同一张表中
    titleFile = os.path.join(os.getcwd(),'read','汇总表头.xlsx')
    titleList = read_excel(titleFile,0)

    # 生成合并表
    #5
    fiveWork = xlwt.Workbook()
    fiveSheet = fiveWork.add_sheet("数据")
    write_rows(fiveSheet,titleList[0],0)
    write_sheet(fiveSheet,fiveData,[],1)
    fiveWork.save(fiveFile)
    #4
    fourWork = xlwt.Workbook()
    fourSheet = fourWork.add_sheet("数据")
    write_rows(fourSheet, titleList[0], 0)
    write_sheet(fourSheet, fourData,[],1)
    fourWork.save(fourFile)
    #2
    twoWork = xlwt.Workbook()
    twoSheet = twoWork.add_sheet("数据")
    write_rows(twoSheet, titleList[0], 0)
    write_sheet(twoSheet, twoData, col2[0], 1)
    twoWork.save(file2Out)


