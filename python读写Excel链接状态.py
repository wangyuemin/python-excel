#-*- coding: utf8 -*-
import xlrd
from xlrd import open_workbook
from xlutils.copy import copy
import requests
ln=[]

fname = "D:\\1.xlsx"
bk = xlrd.open_workbook(fname)
shxrange = range(bk.nsheets)
wb = copy(bk)
try:
 sh = bk.sheet_by_name("rakuten(305)")
 data = xlrd.open_workbook(fname)
 ws = wb.get_sheet("rakuten(305)")
except:
 print "no sheet in %s named Sheet1" % fname
#获取行数
nrows = sh.nrows
#获取列数
ncols = sh.ncols
print "nrows %d, ncols %d" % (nrows,ncols)
#print sh.cell_value(2,3)
#获取第一行第一列数据

for i in range(2,nrows):
    try:
        cell_value = sh.cell_value(i,4)
        code=requests.get(cell_value,timeout=20).status_code
        print i
        print cell_value
        print code
        ws.write(i, 6, code)
        wb.save("D:\\1.xlsx")
    except:
        pass



'''
row_list = []
#获取各行数据
for i in range(1,nrows):
 row_data = sh.row_values(i)
 row_list.append(row_data)
'''
