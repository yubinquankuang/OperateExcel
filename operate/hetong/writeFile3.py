import os
import re
import xlwt
import xlrd
from operate.hetong.paramData import col4, col5, readf, writef
from base.baseOperate import *


def getDuplicateByGoodsMoney(datas, contract, goods, money):
    '''
    对数据进行筛选
    合同号相同的根据goods列判断是否有03开头的，如果存在03开头的编号，删除其他号码开头的数据
    如果没有03的合同，则挑选其余合同中money最大的数据
    datas: 源数据
    contract: 合同号列，base=0
    goods: 货物编码列，base=0
    money: 总金额列, base=0
    '''
    duplicate_counts = 0
    index = 0
    data_len = len(datas)
    max_counts = 0
    work_datas = []
    err_datas = []
    wait_deal = []

    for data in datas:
        index += 1
        if not wait_deal:
            wait_deal.append(data)
        else:
            if data[0] == wait_deal[0][0]:
                wait_deal.append(data)
                if index == data_len:
                    duplicate_counts += 1
                    result = delGoodsMoney(wait_deal, goods, money)
                    print('len', len(result[0]), len(result[1]))
                    work_datas.extend(result[0])
                    err_datas.extend(result[1])
                    max_counts += result[2]
            else:
                if len(wait_deal) > 1:
                    duplicate_counts += 1
                    result = delGoodsMoney(wait_deal, goods, money)
                    print('len',len(result[0]),len(result[1]))
                    work_datas.extend(result[0])
                    err_datas.extend(result[1])
                    max_counts += result[2]
                else:
                    work_datas.extend(wait_deal)
                wait_deal = []
                wait_deal.append(data)
                if index == data_len:
                    work_datas.extend(wait_deal)

    print("重复数据个数：", duplicate_counts)
    print("max数：", max_counts)
    return [work_datas, err_datas]


def delGoodsMoney(datas, goods, money):
    """
    处理合同号相同的数据列表：
    datas: 合同号相同的数据列表
    goods: 货物编号列下标
    money: 总金额编号列下标
    """
    work = []
    err = []
    max_count =0
    goodsResult = False # 根据货物还是根据金钱执行

    for data in datas:
        temp = re.findall(r'^03\d+', data[int(goods)])
        if temp:
            goodsResult = True
            break
    if goodsResult:
        # 根据03筛选
        for data in datas:
            temp = re.findall(r'^03\d+', data[goods])
            if temp:
                work.append(data)
            else:
                err.append(data)
    else:
        max_count = 1
        sort_data = sorted(datas, key=lambda data: data[money],reverse= True)
        work.append(sort_data[0])
        err.extend(sort_data[1:])
    return [work,err,max_count]


if __name__ == "__main__":
    dataFile1 = readf('5.xls')
    data1 = read_excel(dataFile1)

    # 获取去重数据
    workdatas, errdatas = getDuplicateByGoodsMoney(data1, 0, 22, 27)
    print("work: ", len(workdatas), "err:", len(errdatas))

    # 将数据写入对应的表格
    workFile1 = writef('5d.xls')
    errFile1 = writef('5e.xls')
    sheet_name = "数据"
    title = read_excel(dataFile1,0,2)[0]

    workBook1 = xlwt.Workbook()
    errBook1 = xlwt.Workbook()

    workSheet = workBook1.add_sheet(sheet_name)
    errSheet = errBook1.add_sheet(sheet_name)

    print(workdatas[-1])
    print(errdatas[-1])

    write_rows(workSheet, title, 0)
    write_sheet(workSheet, workdatas,[], 1)
    write_rows(errSheet, title, 0)
    write_sheet(errSheet, errdatas,[], 1)

    errBook1.save(errFile1)
    print("work save")
    workBook1.save(workFile1)
    print('5表数据个数：', len(data1))

    # ++++++++++++++++++++++++++++++
    dataFile1 = readf('4-2N4com.xlsx')
    data1 = read_excel(dataFile1)

    # 获取去重数据
    workdatas, errdatas = getDuplicateByGoodsMoney(data1, 0, 22, 27)
    print("work: ", len(workdatas), "err:", len(errdatas))

    # 将数据写入对应的表格
    workFile1 = writef('4-2N4d.xls')
    errFile1 = writef('4-2N4e.xls')
    sheet_name = "数据"
    title = read_excel(dataFile1, 0, 2)[0]

    workBook1 = xlwt.Workbook()
    errBook1 = xlwt.Workbook()

    workSheet = workBook1.add_sheet(sheet_name)
    errSheet = errBook1.add_sheet(sheet_name)

    print(workdatas[-1])
    print(errdatas[-1])

    write_rows(workSheet, title, 0)
    write_sheet(workSheet, workdatas, [], 1)
    write_rows(errSheet, title, 0)
    write_sheet(errSheet, errdatas, [], 1)

    errBook1.save(errFile1)
    print("work save")
    workBook1.save(workFile1)
    print('4-2N4表数据个数：', len(data1))
