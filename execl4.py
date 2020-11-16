#author: sven
#date: 2020/10/26
#

#!/usr/bin/python
# -*- coding: UTF-8 -*-

import xlrd
import sys
import os

keyCloumName = u'Keyname开发变量名'
keyCloumName2 = u'Keyname'
        

path = sys.argv[1]
workbook = xlrd.open_workbook(path)
  
 # 输出Excel文件中所有sheet的名字
print(workbook.sheet_names())

#保存文件
def write_file(file_path, content):
    with open(file_path, "a+") as f:
        f.writelines(content)

#寻找key对应的列
#table:对应的表
#返回所在列
def getKeyNameCol(table):
    for i in range(table.ncols):
        title = table.col(i)[0].value
        if (title == keyCloumName or title == keyCloumName2):
            return i


#判断文件夹是否存在，并新建
def make_path(p):
    path = "%s/%s"%((os.path.split(os.path.realpath(__file__))[0]), p)
    if not os.path.exists(path):
        os.makedirs(path)
    return path



#将key，value组队
#table:对应的表
#keyCol:key对应在哪列
def formatKeyValue(table, keyCol):
    rowNum = table.nrows  # table行数
    colNum = table.ncols  # table列数
    for m in range(table.ncols):
        if len(table.col(m)[keyCol].value) > 0:
            title = table.col(m)[0].value
            print ("---------------- "+ title +" ----------------")
            for n in range(rowNum):
                key = table.col(keyCol)[n].value
                if key.startswith(' '):
                    print ("xxxx---key中有空格前缀" + key)
                if key.endswith(' '):
                    print ("xxxx---key中有空格后缀" + key)

                while key.startswith(' '):
                    key = key[1:]
                while key.endswith(' '):
                    key = key[:len(key)-1]
                value = table.col(m)[n].value
                if value.startswith(' '):
                    print ("xxxx---value中有空格前缀" + value)
                if value.endswith(' '):
                    print ("xxxx---value中有空格后缀" + value)

                while value.startswith(' '):
                    value = value[1:]
                while value.endswith(' '):
                    value = value[:len(value)-1]

                if value.find('\\"') >= 0:
                    pass
                elif value.find('\"') >= 0:
                    value = value.replace('\"', '\\"')
                else:
                    pass

                if len(key) > 0 and key != value and key != keyCloumName:
                    str_ios = "\""+ key + "\"" +" = " +"\""+ value + "\"" + ";" + "\n" 
                    str_android = "<string name=" + "\""+ key + "\">" + value + "</string>" + "\n"
                    print ("%s ------ %s"%(key, value))
                    path_ios = "%s/%s.txt"%(make_path("txt/ios"),title)
                    write_file(path_ios, str_ios)

                    path_android = "%s/%s.txt"%(make_path("txt/android"),title)
                    write_file(path_android, str_android)
            print("\n")

if __name__ == '__main__':
    if len(sys.argv)!=2:
        print("命令格式错误，应该为：python3"+" "+__file__+" "+"‘输入文件路径’")
        exit()
        
    for k in range(len(workbook.sheet_names())):
        eachTable = workbook.sheets()[k]
        key_col = getKeyNameCol(eachTable)
        if key_col >= 0:
            formatKeyValue(eachTable, key_col)





	







	
