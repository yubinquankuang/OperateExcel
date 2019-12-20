import os
import re
import xlrd
import xlwt
from operate.salesMoney.paramData import *
from base.baseOperate import *


def fillContract(datas):
    """
    保留存货编号前面的部分同时
    """
    cur_contract = ""
    new_datas = []
    for data in datas:
        if data[contractNum]:
            cur_contract = data[contractNum]
        else:
            data[contractNum] = cur_contract
        temp = re.findall(r"^(03|02|01)", str(data[stockNum]))
        if temp:
            data[stockNum] = "清空"
            data[stockNum+1:] = ''
        new_datas.append(data)
    return new_datas

def matchContract(numlist, datas):
    """
    根据numlist中的合同号匹配datas中的数据，返回匹配到的数据
    numlist： 编号列表
    datas: 数据
    """
    new_datas = []
    for data in datas:
        for num in numlist:
            if num in data[0]:
                new_datas.append(data)
                break
    return new_datas

def delBlankData(datas):
    """
    删除空白数据，筛选依据是第10列为清空，且8 9 为空
    """
    err_datas = [datas[0],datas[1]]
    new_datas = [datas[0],datas[1]]
    for data in datas[2:]:
        if ((data[9] != "清空") or (data[8] != "")) and (data[10] or data[11] or data[12] or data[8]):
            new_datas.append(data)
        else:
            err_datas.append(data)
    return [new_datas,err_datas]

def mergeData(datas):
    """
    根据一定的规则进行合并
    """
    err_datas = [datas[0], datas[1]]
    new_datas = [datas[0], datas[1]]
    temp = []
    for data in datas[2:]:
        pass


if __name__ == "__main__":
    # 读取原始表格数据
    fileName1 = readf('sales.xlsx')
    data1 = read_excel(fileName1,0,0,3)

    # 填充合同号 同时输出文件
    data2 = fillContract(data1)
    fileName2 = writef("sales1.xlsx")
    write_file(fileName2,"数据",data2)

    # 获取合同号列表
    saleNumFile = writef("salesNum.xlsx")
    saleNum = getComNums(fileName2, saleNumFile) # 合同号 对应行
    write_file(saleNumFile, "合同列表", saleNum)

    # 匹配文件获取对应列的数据
    numlist = [x[0] for x in saleNum]
    data3 = read_excel(fileName1,1,0,2,readCols)
    data4 = matchContract(numlist, data3)
    print(data4)

    # 生成文件
    fileName3 = writef("matchFile.xlsx")
    work2 = xlwt.Workbook()
    sheet2 = work2.add_sheet("finalData")
    write_sheet(sheet2, data2)
    new_start = len(data2)
    write_sheet(sheet2,data4,writeCols,new_start)
    work2.save(fileName3)

    # 处理数据, 删除全部为空的数据
    fileName4 = readf("matchFile.xlsx")
    data5 = read_excel(fileName4,0) # 第一行为标题
    print(data5[0])
    data5new, data5err = delBlankData(data5)
    fileName5 = writef("newMatch.xlsx")
    fileName6 = writef("errMatch.xlsx")
    write_file(fileName5,"数据",data5new)
    write_file(fileName6, "数据",data5err)

    # 将处理过的数据合并
    data6 = read_excel(fileName5,0)


