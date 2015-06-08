#!/usr/bin/python
# -*- coding: utf_8 -*-

__author__ = '岸城'

import xlrd
import glob


import sqlite3

# 打开数据库文件
ExcelComp_db = sqlite3.connect('ExcelComp.db')
ExcelComp_db.text_factory = str
cursor = ExcelComp_db.cursor()

# 建表
cursor.execute('DROP TABLE IF EXISTS ExcelComp')
cursor.execute('DROP TABLE IF EXISTS ExcelTempData')
cursor.execute('CREATE TABLE ExcelComp (id INTEGER PRIMARY KEY, file_name varchar(128),company_name varchar(256),odds_win varchar(16),odds_equal varchar(16),odds_lose varchar(16))')
cursor.execute('CREATE TABLE ExcelTempData (id INTEGER PRIMARY KEY, file_name varchar(256),company_name varchar(256),odds_win varchar(16),odds_equal varchar(16),odds_lose varchar(16))')

# 打开 device 输入 Excel 文件

for filename in glob.glob(r'*.xls*'):
    print filename
    device_workbook = xlrd.open_workbook(filename)
    for worksheet_name in device_workbook.sheet_names():
        device_sheet = device_workbook.sheet_by_name(worksheet_name)

    #读出xls内容并写数据库
    for row in range(6, device_sheet.nrows):
       company_name = device_sheet.cell(row, 0).value
       odds_win = device_sheet.cell(row, 11).value
       odds_equal = device_sheet.cell(row, 12).value
       odds_lose = device_sheet.cell(row, 13).value
       cursor.execute('INSERT INTO ExcelComp (file_name,company_name,odds_win,odds_equal,odds_lose) VALUES (?,?,?,?,?)',
                      (filename.decode('gbk').encode("utf-8"),company_name,odds_win,odds_equal,odds_lose))

    ExcelComp_db.commit()


# 按需查询数据库
cursor.execute('SELECT file_name,company_name,odds_win,odds_equal,odds_lose '
               'FROM ExcelComp AS a ')

rows = cursor.fetchall()
brows = list(rows) #复制数组


# 建立一个temp数组便于存储xls文件名
temparr=[]
i=0
for row in rows:
    temparr.append(row[0])
    i=i+1
    for num in range(i,len(brows)):
        if row[0]!=brows[num][0] and \
            row[1]==brows[num][1] and \
            row[2]==brows[num][2] and \
            row[3]==brows[num][3] and \
            row[4]==brows[num][4]:
            temparr.append(brows[num][0])

    # for brow in brows:
    #     if row[0]!=brow[0] and \
    #     row[1]==brow[1] and \
    #     row[2]==brow[2] and \
    #     row[3]==brow[3] and \
    #     row[4]==brow[4]:
    #         temparr.append(brow[0])


    if len(temparr)>1:
        #写公司名赔率插入临时表
        cursor.execute('INSERT INTO ExcelTempData (file_name,company_name,odds_win,odds_equal,odds_lose) VALUES (?,?,?,?,?)',
                      (';'.join(temparr),row[1],row[2],row[3],row[4]))
    #清空数组
    del temparr[:]
#执行写数据库操作
ExcelComp_db.commit()


# 建立一个文本文件用作输出
output = open('ExcelComp.txt', 'w')
# 按需查询数据库
cursor.execute('select file_name,company_name,odds_win,odds_equal,odds_lose  from ExcelTempData order by file_name  ')
rows = cursor.fetchall()
tempstr=""
cname=""
for row in rows:

    if tempstr!=row[0]: #当前文件名不等于上一个则写入
        output.write('\n')
        tempstr=str(row[0])
        t=tempstr.split(';')
        if len(t)>0:
            for x in t:
                output.write('\n'+x) #写文件名
        #纪录第一个公司名和数据
        cname='\n        '+row[1]+","+row[2]+","+row[3]+","+row[4]
        print cname
    else:
        if cname!="":
            output.write(cname)
            cname=""
        output.write('\n        '+row[1]+","+row[2]+","+row[3]+","+row[4])
    if cname!="":
        output.write(cname)
        cname=""

#关闭所有
output.close()
ExcelComp_db.close()












