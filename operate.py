from data import *
from ReadXls import read_excel
from WriteXls import *



if __name__ == "__main__":
    data = read_excel(read_path,1,4155)
    write_excel(write_path,err_path,data)
    data1 = read_excel(write_path,0,4153)
    write_final_xls(final_path,data1)