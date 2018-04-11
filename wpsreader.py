#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__author__ = 'Hunter'

import win32com.client
from express import Parcel
import CONST


class WPS(object):
    def __init__(self, xls_path):
        self.xls_path = xls_path
        self.dict_all = {}

    @property
    def read_xls(self):
        parcels = []
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = -1
        my_book = excel.Workbooks.Open(self.xls_path)
        my_sheet = my_book.Worksheets("LIST")
        row = my_sheet.usedrange.rows.count
        # col = mySheet.usedrange.columns.count
        i = CONST.WPS_START_ROW
        cnt = 1
        while i < row:
            num = my_sheet.Cells(i, CONST.WPS_NAME_ROW).Value
            name = my_sheet.Cells(i, CONST.WPS_NAME_ROW).Value
            tel = my_sheet.Cells(i, CONST.WPS_TEL_ROW).Value
            address = my_sheet.Cells(i, CONST.WPS_ADDRESS_ROW).Value
            if num is None:
                print(i, '无效单号')
            elif name is None:
                print(i, '无效收件人')
            elif tel is None:
                print(i, '无效收件人号码')
            elif address is None:
                print(i, '无效收件人地址')
            else:
                parcel = Parcel(num, name, tel, address)
                parcels.append(parcel)
                self.dict_all[cnt] = parcel
                cnt += 1
            # names.append(str(my_sheet.Cells(i, 2).Value))
            i += 1
        return parcels

    @property
    def filter_parcels(self, _parcel_list):
        name_list = {}
        for _parcel in _parcel_list:
            if name_list.get(_parcel.rec_name) is None:
                # print(_parcel.rec_name)
                name_list[_parcel.rec_name] = _parcel.num
            else:
                print(_parcel.rec_name)
        return True

wps = WPS(r'C:\Users\Hunter\PycharmProjects\starex\快件核对_0421150005.xls')
parcel_list = wps.read_xls
# wps.filter_parcels(parcel_list)
print(wps.dict_all)
# filter_parcel()
# for _parcel in parcel_list:
#     print(':', _parcel.num)