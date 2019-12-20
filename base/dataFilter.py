
def getDuplicate(datas, cols):
    """
    判断前首相将表头和第一行数据加入输出数列
    """
    if len(cols) == 0 or len(datas) < 2:
        return datas
    new_datas = [datas[0],datas[1]]
    for data in datas[2:]:
        result = False
        for col in cols:
            if data[col] != new_datas[-1][col]:
                result = True
                break
        if result:
            new_datas.append(data)

    return new_datas




if __name__ == "__main__":
    datas = [[12,"1",333],[12,"2",33],[12,'1',333],[12,"2",33],[12,"2",33]]
    print(getDuplicate(datas,[0,2]))