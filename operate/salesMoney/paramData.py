import os


# func
def readf(name):
    return os.path.join(os.getcwd(), 'read', name)

def writef(name):
    return os.path.join(os.getcwd(), 'write', name)


# data
stockNum = 9
contractNum = 0
readCols = [0,8,9,10,11,12,13,14,16]
writeCols = [0,10,9,28,11,12,13,14,19]