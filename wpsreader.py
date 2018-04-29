#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__author__ = 'Hunter'

import win32com.client
from express import Parcel
import CONST


def select_repeat(a, b):
    if a.rec_tel == b.rec_tel:
        print("xxx", a.num, b.num)
        return b.num


class WPS(object):
    def __init__(self, xls_path):
        self.xls_path = xls_path
        self.dict_all = {}  # <id,包裹>
        self.result_list = []  # 最终得到的过滤后的 id_List
        self.repeat_list = []  # 重复件id_List

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
            num = my_sheet.Cells(i, CONST.WPS_NUM_ROW).Value
            name = my_sheet.Cells(i, CONST.WPS_NAME_ROW).Value
            tel = my_sheet.Cells(i, CONST.WPS_TEL_ROW).Value
            address = my_sheet.Cells(i, CONST.WPS_ADDRESS_ROW).Value
            if num is None:
                print(i, 'Error code.')
            elif name is None:
                print(i, 'Error receptor.')
            elif tel is None:
                print(i, 'Error number.')
            elif address is None:
                print(i, 'Error address.')
            else:
                parcel = Parcel(i, num, name, tel, address)
                parcels.append(parcel)
                self.dict_all[cnt] = parcel
                cnt += 1
            # names.append(str(my_sheet.Cells(i, 2).Value))
            i += 1
            self.result_list = range(1, self.dict_all.__len__())
        return parcels

    # 根据姓名过滤
    def filter_name(self):
        tep_namelist = {}  # 临时姓名字典 <姓名,包裹>
        rep_namelist = {}  # 同名包裹字典 <姓名+tel，同名包裹id的list>
        for _parcel in self.dict_all.values():
            pre_parcel = tep_namelist.get(_parcel.rec_name)  # 根据当前的包裹的姓名从临时tep_namelist字典里查找前一个同名的包裹
            if pre_parcel is None:  # 如果没有同名的包裹，则将当前的包裹放入临时的tep_namelist字典
                # print(_parcel.rec_name)
                tep_namelist[_parcel.rec_name] = _parcel
            else:  # 有同名的id ，则看手机是否相同：如果都相同，则先放入同名字典
                if _parcel.rec_tel == pre_parcel.rec_tel:
                    tep_key = _parcel.rec_name+_parcel.rec_tel
                    tep_list = rep_namelist.get(tep_key)
                    if tep_list is None:
                        rep_namelist[tep_key] = [pre_parcel.ser_id]
                    else:
                        tep_list.append(_parcel.ser_id)
                    # print(_parcel.rec_name)
                    # tmp_id = select_repeat(pre_parcel, _parcel)
                    # if tmp_id is not 0:
                    #     self.repeat_list.append(tmp_id)
        print(rep_namelist)


wps = WPS(r'C:\Users\Hunter\PycharmProjects\starex\快件核对_0417131011.xls')
parcel_list = wps.read_xls()
wps.filter_name()
print(wps.repeat_list.__len__())


# filter_parcel()
# for _parcel in parcel_list:
#     print(':', _parcel.num)
