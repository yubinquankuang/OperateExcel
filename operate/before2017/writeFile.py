import re
import os
import xlwt
import xlrd
from base.baseOperate import *
from operate.before2017.paramData import *


def duplicateData(datas):
    """
    处理重复数据，根据指定的列进行去重操作
    data: 要去重的数据
    """
    pre_invoice = 0
    new_datas = []
    for data in datas:
        print("invoice",data[invoiceIndex])
        if data[invoiceIndex] == pre_invoice:
            new_contract = " " + data[contractIndex]
            new_datas[-1][contractIndex] += new_contract
        else:
            if new_datas:
                print("contract: ",new_datas[-1][contractIndex])
            pre_invoice = data[invoiceIndex]
            new_datas.append(data)
    return new_datas




if __name__ == "__main__":
    fileData1 = read_excel(fileIn1,0,0,sheetIndex)
    # 数据去重
    fileData = duplicateData(fileData1)
    # 生成文件
    outWork = xlwt.Workbook()
    outSheet = outWork.add_sheet(sheetOutName)
    write_sheet(outSheet, fileData)
    outWork.save(fileOut1)

    print(len(fileData1))