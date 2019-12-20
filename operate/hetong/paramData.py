import os

cols = [0, 3, 2, 3, 3, 3, 0]

cols7 = [7, 10, 9, 10 , 10 , 10 , 7]

# 去重列表
duplicate = [
    [],
    [2, 3, 6, 12],
    [2, 5, 7],
    [3, 13, 14],
    [3, 13, 14],
    [3, 13, 14],
    [3, 13, 14]
    # [3],
    # [2],
    # [3],
    # [3],
    # [3],
    # [3]
]

# 4
col4 = [
    [21,27],
    [-1,-1,27,16,5,-1,-1,9,-1,4,-1,21,-1,15,-1,-1],
    [-1,-1,16,5,-1,-1,9,27,-1,4,-1,21,-1,15,-1,-1],
    [16,19,20,-1,-1,9,18,5,21,22,23,24,25,26,27,28,29,-1,10,30,31],
    [-1,19,20,16,-1,9,18,5,21,22,23,24,25,26,27,28,29,-1,10,30,31,32],
    [16,19,20,-1,-1,9,18,5,21,22,23,24,25,26,27,28,29],
    [16,19,20,-1,-1,9,18,5,21,22,23,24,25,26,27,28,29,-1,10,30,31]
]

# 5
col5 = [
    [0,5,8,10,11,12,18,21,27],
    [0, 5, 8, 10, 11, 12, 18,-1,-1,27,16,-1,-1,-1,9,-1,4,-1,21,-1,15,-1,-1],
    [0, 5, 8, 10, 11, 12, 18,-1,-1,16,-1,-1,-1,9,27,-1,4,-1,21,-1,15,-1,-1],
    [0, 5, 8, 10, 11, 12, 18,16,19,20,-1,-1,9,-1,-1,21,22,23,24,25,26,27,28,29,-1,-1,30,31],
    [0, 5, 8, 10, 11, 12, 18,-1,19,20,16,-1,9,-1,-1,21,22,23,24,25,26,27,28,29,-1,-1,30,31,32],
    [0, 5, 8, 10, 11, 12, 18,16,19,20,-1,-1,9,-1,-1,21,22,23,24,25,26,27,28,29],
    [0, 5, 8, 10, 11, 12, 18,16,19,20,-1,-1,9,-1,-1,21,22,23,24,25,26,27,28,29,-1,-1,30,31]
]

# 2
col2 = [
    [0,5,9,10,11,12,18]
]

def readf(name):
    return os.path.join(os.getcwd(), 'read', name)

def writef(name):
    return os.path.join(os.getcwd(), 'write', name)