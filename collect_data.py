
import os
import xlrd
from xlutils.copy import copy
 
 
def readdir(rootdir):
    excel = r'...\mix.xls'
    try:
        rdx = xlrd.open_workbook(excel, formatting_info=True)   #打开Excel，并保留原格式
    except:
        print "no excel in %s " % excel
    wtx = copy(rdx)                                             #复制为可读写的wtx
    sheet = wtx.get_sheet(0)
    i = 0
    for parent, dirnames, filenames in os.walk(rootdir):
        print parent                                            #查看文件的父目录    
        for filename in filenames:                              #将文件夹下所有文件写入Excel，每个文件
            j = 0                                               #为一行，文件中每一行占一个单元格
            abspath = os.path.join(parent, filename)
            f = open(abspath, 'r')
            for lines in f.readlines():
                sheet.write(i, j, lines)
                j += 1
            i += 1
    print("write new information successfully")
    wtx.save(excel)
    print ("save the information successfully!")
