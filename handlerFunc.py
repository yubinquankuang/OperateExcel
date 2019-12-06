import re

def delDuplicate(data):
    """
    合同信息
    """
    new_data = []
    count = 0
    count1 = 0
    for row in data:
        if count == 0 or row[0] != new_data[-1][0]:
            new_data.append(row)
            count += 1
        else:
            count1 += 1
    print(len(new_data),count1)
    return new_data

def delTestEquipment(data):
    """
    合同设备调试处理函数
    """
    new_data = []
    new_data.append(data[0])
    for row in data[1:]:
        if row[25] == '没找到' or row[25] == '没法找':
            row[25] = ''
        if row[24] == '无需调试' or row[22] == "无03成品":
            continue
        if row[22] == row[23]:
            continue
        if len(re.findall('\d{5,}', row[22])) > 0:
            row[22] = re.findall('\d{5,}',row[22])[0]
        else:
            continue
        new_data.append(row)
    return new_data

def delInvoice(data):
    """
    合同开票凭证处理函数
    """
    new_data = []
    new_data.append(data[0])
    for row in data[1:]:
        if isinstance(row[12],float) and row[12] != 0:
            new_data.append(row)
    print(len(new_data),len(data))
    return new_data

def delMakeSureIncome(data):
    '''
    合同确认收入
    '''
    new_data = []
    new_data.append(data[0])
    no_tax = 28
    tax = 29
    for row in data[1:]:
        if (row[no_tax] != 0 and row[no_tax] != '' and row[no_tax] != '0') or (row[tax] != 0 and row[tax] != '' and row[tax] != '0'):
            new_data.append(row)
    return new_data

def delCost(data):
    '''
    合同成本转结
    '''
    new_data = []
    new_data.append(data[0])
    cost = 35
    for row in data[1:]:
        if row[cost] != 0 and row[cost] != '' and row[cost] != '0':
            new_data.append(row)
        else:
            print(row[cost])
    return new_data