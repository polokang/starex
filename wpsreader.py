#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__author__ = 'Hunter'

import win32com.client
from express import Parcel

class WPS(object):
    def __init__(self, xls_path):
        self.xls_path = xls_path

    @property
    def read_xls(self):
        parcels = []
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = -1
        my_book = excel.Workbooks.Open(self.xls_path)
        my_sheet = my_book.Worksheets("LIST")
        row = my_sheet.usedrange.rows.count
        # col = mySheet.usedrange.columns.count
        i = 2
        while i <= row:
            num = my_sheet.Cells(i, 5).Value
            name = my_sheet.Cells(i, 7).Value
            tel = my_sheet.Cells(i, 12).Value
            address = my_sheet.Cells(i, 10).Value
            if num is None:
                print('无效单号')
            elif name is None:
                print('无效收件人')
            elif tel is None:
                print('无效收件人号码')
            elif address is None:
                print('无效收件人地址')
            else:
                parcel = Parcel(num, name, tel, address)
                parcels.append(parcel)
            # names.append(str(my_sheet.Cells(i, 2).Value))
            i += 1
        return parcels


wps = WPS(r'C:\Users\Hunter\PycharmProjects\starex\BNE西安发货数据P48.xls')
parcel_list = wps.read_xls
for _parcel in parcel_list:
    print(':', _parcel.rec_address)
