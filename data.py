'''
参数模块
'''
import os
from handlerFunc import delInvoice, delTestEquipment, delDuplicate, delCost, delMakeSureIncome

COLUMS = [
    {
        'name': '合同信息',
        'cols': [0,4,7],
        'func': delDuplicate
    },
{
        'name': '合同凭证',
        'cols': [0,1,2,3],
        'func': None
    },
{
        'name': '合同开票凭证',
        'cols': [0,1,2,3,6,11,12],
        'func': delInvoice
    },
{
        'name': '合同设备调试情况',
        'cols': [0,9,22,23,24,25],
        'func': delTestEquipment
    },
{
        'name': '合同确定收入',
        'cols': [0,1,2,3,6,28,29],
        'func': delMakeSureIncome
    },
{
        'name': '合同成本转结',
        'cols': [0,1,2,3,6,35],
        'func': delCost
    }
]

mydir = os.getcwd()
write_file_name = "duplicate_remove.xls"
read_file_name = "2017.xlsx"
err_file_name = 'err.xls'
final_name = '最终整理表.xlsx'
out_path = os.path.join(mydir,'data','write')
read_path = os.path.join(mydir,'data','read')
write_path = os.path.join(out_path,write_file_name)
err_path = os.path.join(out_path,err_file_name)
read_path = os.path.join(read_path,read_file_name)
final_path = os.path.join(out_path,final_name)