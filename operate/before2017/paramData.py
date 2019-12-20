import os


# func
def readf(name):
    return os.path.join(os.getcwd(), 'read', name)

def writef(name):
    return os.path.join(os.getcwd(), 'write', name)


# data
sheetIndex = 8
fileIn1 = readf("2017Invoice.xlsx")
contractIndex = 5
invoiceIndex = 1
sheetOutName = "17年前发票"
fileOut1 = writef("2017before.xlsx")

a = ['34','34','3433']
print(" ".join(a))