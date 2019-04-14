#!/usr/bin/env python3
# -*- coding: utf-8 -*-
__author__ = 'Hunter'

import win32com.client
import mysql.connector
conn = mysql.connector.connect(host='150.109.54.77', port='3306', user='starex', password='1111', database='starex')
# conn = mysql.connector.connect(host='localhost', port='3306', user='root', password='123456', database='starex')

cursor = conn.cursor()
# cursor.execute('select * from inventorytype where id = %s', ('1',))
# values = cursor.fetchall()
# print('->',values)


excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = -1
my_book = excel.Workbooks.Open(r'C:\Users\Hunter\Desktop\耗材记录.xlsx')
my_sheet = my_book.Worksheets("18年7月")
row = my_sheet.usedrange.rows.count
i = 5
# i = 43
message = ''

while i < row:
    date = my_sheet.Cells(i, 2).Value
    sql = 'insert into inventory (outdate,inventory,cnt,opname,agent) values(%s, %s, %s, %s, %s)'
    if date is None:
        print('结束')
        break;
    else:
        message = date + ':'
        op = my_sheet.Cells(i, 12).Value
        agent = my_sheet.Cells(i, 13).Value
        if isinstance(agent, str) is not True:
            agent = str(agent)
            agent= agent.split('.')[0]
        else:
            agent = agent.capitalize()
        # sql = 'insert into inventory (id,outdate,inventory,cnt,opname,agent)'
        # param = [('%d ')]
        print(agent)
        _6 = my_sheet.Cells(i, 3).Value
        _4 = my_sheet.Cells(i, 4).Value
        _3 = my_sheet.Cells(i, 5).Value
        _2 = my_sheet.Cells(i, 6).Value
        _1 = my_sheet.Cells(i, 7).Value
        tape = my_sheet.Cells(i, 8).Value
        pop = my_sheet.Cells(i, 9).Value
        if _6 is not None:
            message = message + ' ' + str(_6)
            param = (date, 6, _6, op, agent)
            cursor.execute(sql, param)
        if _4 is not None:
            message = message + ' ' + str(_4)
            param = (date, 4, _4, op, agent)
            cursor.execute(sql, param)
        if _3 is not None:
            message = message + ' ' + str(_3)
            param = (date, 3, _3, op, agent)
            cursor.execute(sql, param)
        if _2 is not None:
            message = message + ' ' + str(_2)
            param = (date, 2, _2, op, agent)
            cursor.execute(sql, param)
        if _1 is not None:
            message = message + ' ' + str(_1)
            param = (date, 1, _1, op, agent)
            cursor.execute(sql, param)
        if tape is not None:
            message = message + ' ' + str(tape)
            param = (date, 7, tape, op, agent)
            cursor.execute(sql, param)
        if pop is not None:
            message = message + ' ' + str(pop)
            param = (date, 8, pop, op, agent)
            cursor.execute(sql, param)

    print(str(i),message + '->' + op + ':', agent)
    i += 1
    message = ''

# sql = 'insert into inventory (id,outdate,inventory,cnt,opname,agent)'
# param = [('%d ')]
conn.commit()
conn.close()
