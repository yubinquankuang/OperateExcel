import os
import re
import xlrd
import xlwt
from operate.salesMoney1.paramData import *
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
        flag = False
        index = -1
        for d in data:
            index += 1
            if index in delCols:
                continue
            if d:
                flag = True
                break

        if flag:
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
    data_len = len(datas)
    temp = []
    index = 1
    for data in datas[2:]:
        index += 1
        if temp:
            if data[0] == temp[0][0]:
                temp.append(data)
            else:
                if len(temp) > 1:
                    temp = delConstract(temp)
                    err_datas.extend(temp[1])
                    temp = temp[0]
                new_datas.extend(temp)
                temp = []
                temp.append(data)
            if index == data_len - 1:
                if len(temp) > 1:
                    temp = delConstract(temp)
                    err_datas.extend(temp[1])
                    temp = temp[0]
                new_datas.extend(temp)
        else:
            temp.append(data)
    return [new_datas,err_datas]

def delConstract(datas):
    """
    处理合同号相同的数据
    :param datas:
    :return: new_datas
    """
    pre = []
    add = []
    new_datas = []
    err_datas = []
    index = -1
    # 获取要处理的编号列表，pre为要处理的编号，add为要添加的编号
    for data in datas:
        index += 1
        check1 = False
        for num in proCols:
            if data[num]:
                check1 = True
                break
        if check1:
            check2 = True
            for num in afterCols:
                if data[num]:
                    check2  = False
                    break
            if check2:
                pre.append(index)
        else:
            add.append(index)

    delete = []
    index = -1
    addlen = len(add)
    prelen = len(pre)
    maxlen = addlen
    if addlen >= prelen:
        maxlen = prelen
    if add and pre:
        for ad in add:
            index += 1
            if index < maxlen:
                p = pre[index]
                a = add[index]
                changeList(datas[p],datas[a])
                delete.append(add[index])
        # new_datas = []
        index = -1
        for data in datas:
            index +=1
            if index in delete:
                err_datas.append(data)
                continue
            else:
                new_datas.append(data)
            print(len(err_datas))

        return [new_datas, err_datas]
    else:
        return [datas,[]]

def changeList(origin, change):
    """
    替换指定列的数据
    :param origin:
    :param change:
    :return:
    """
    origin[9:] = change[9:]
    return origin



if __name__ == "__main__":
    filename1 = readf("salesMoney.xlsx")
    data1 = read_excel(filename1,0,0,3)
    data1Fill = fillContract(data1)
    filename2 = writef("newSalesMoney.xlsx")
    write_file(filename2,"data",data1Fill)


    # 数据分析,获取参照列表
    data1NumsFile = writef("oneNum.xlsx")  # 注意必须是表格中的第一张sheet
    data1Nums = getComNums(filename2, data1NumsFile, 0)
    numlist1 = list(set([x[0] for x in data1Nums[1:]]))

    # 数据匹配并加入
    data2 = read_excel(filename1,1,0,2,readCols)
    data3 = matchContract(numlist1, data2)

    # 将数据写入xlsx文件
    filename3 = writef("combin1.xls")
    work1 = xlwt.Workbook()
    sheet1 = work1.add_sheet("data")
    write_sheet(sheet1,data1)
    write_sheet(sheet1,data3,writeCols,len(data1))
    work1.save(filename3)

    # 去重明显空的数据
    filename5 = readf('combin1.xls')
    data4 = read_excel(filename5, 0)
    data4new,data4err = delBlankData(data4)
    print(len(data4new), len(data4err))

    file4err = writef("file4err.xls")
    file4new = writef("file4new.xls")

    write_file(file4err,"data",data4err)
    write_file(file4new,"data",data4new)

    # 数据合并
    filename4 = writef("final.xls")
    filename5 = writef("finalerr.xls")
    data5 = read_excel(file4new,0)
    print(data5[1])
    print(len(data5))
    data6,data7 = mergeData(data5)
    print(len(data6))
    write_file(filename4,"data",data6)
    write_file(filename5, "data", data7)


