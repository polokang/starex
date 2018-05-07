#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__author__ = 'Hunter'

import time
import win32com.client
from express import Parcel
import CONST

# 返回重复件id,规则如下：
# 1.取时间近的包裹
# 2.取奶粉
# 3.取重量重的
# 4.返回0则表示无法确定
def select_repeat(a, b):
    time_a = time.mktime(time.strptime(a.date, '%Y.%m.%d'))
    time_b = time.mktime(time.strptime(b.date, '%Y.%m.%d'))
    if time_a > time_b:
        return a.ser_id
    elif time_a < time_b:
        return b.ser_id
    else:
        # 判断是否是婴儿奶粉
        if '婴儿奶粉' in a.goods and '婴儿奶粉' in b.goods:
            level_a = a.goods[a.goods.find('段') - 1]
            level_b = b.goods[b.goods.find('段') - 1]
            # print('全是奶粉', a.num, a.goods, '--', b.num, b.goods)
            if level_a <= level_b:
                return a.ser_id
            else:
                return b.ser_id
        else:
            if '婴儿奶粉' in a.goods:
                return a.ser_id
            elif '婴儿奶粉' in b.goods:
                return b.ser_id
            else:
                # 保健品
                if a.weight <= b.weight:
                    return a.ser_id
                else:
                    return b.ser_id
    return 0





class WPS(object):
    def __init__(self, xls_path):
        self.xls_path = xls_path
        self.dict_all = {}  # <id,包裹>
        self.result_list = []  # 最终得到的过滤后的 id_List
        self.repeat_list = []  # 重复件id_List

    def read_xls(self):
        parcels = []
        self.excel = win32com.client.Dispatch('Excel.Application')
        self.excel.Visible = -1
        my_book = self.excel.Workbooks.Open(self.xls_path)
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
            date = my_sheet.Cells(i, CONST.WPS_DATE_ROW).Value
            goods = my_sheet.Cells(i, CONST.WPS_GOODS_ROW).Value
            weight = my_sheet.Cells(i, CONST.WPS_WGT_ROW).Value
            if num is None:
                print(i, 'Error code.')
            elif name is None:
                print(i, 'Error receptor.')
            elif tel is None:
                print(i, 'Error number.')
            elif address is None:
                print(i, 'Error address.')
            else:
                addrcode = str(my_sheet.Cells(i, CONST.WPS_ADDRESS_CODE_ROW).Value)
                parcel = Parcel(cnt, num, name, tel, address)
                parcel.set_customer( my_sheet.Cells(i, CONST.WPS_CUSTOMER_ROW).Value)
                parcel.set_customer_code( my_sheet.Cells(i, CONST.WPS_CUSTOMER_CODE_ROW).Value)
                parcel.set_rec_privince( my_sheet.Cells(i, CONST.WPS_REV_PRINCE_ROW).Value)
                parcel.set_rec_city( my_sheet.Cells(i, CONST.WPS_REV_CITY_ROW).Value)
                parcel.set_address_code(addrcode)
                parcel.set_sender( my_sheet.Cells(i, CONST.WPS_SENDER_NAME_ROW).Value)
                parcel.set_distination( my_sheet.Cells(i, CONST.WPS_DESTINATION_ROW).Value)
                parcel.set_rev_email( my_sheet.Cells(i, CONST.WPS_REV_EMAIL_ROW).Value)
                parcel.set_sender_tel( my_sheet.Cells(i, CONST.WPS_SENDER_TEL_ROW).Value)
                parcel.set_state( my_sheet.Cells(i, CONST.WPS_STATE_ROW).Value)

                parcel.set_date(date)
                parcel.set_goods(goods)
                parcel.set_weight(weight)
                parcels.append(parcel)
                self.dict_all[cnt] = parcel
                self.result_list.append(cnt)  # fill id
                cnt += 1
            # names.append(str(my_sheet.Cells(i, 2).Value))
            i += 1

        return parcels

    # 一.根据姓名过滤
    def filter_name(self):
        tmp_namelist = {}  # 临时姓名字典 <姓名+tel,包裹>
        for _parcel in self.dict_all.values():
            tmp_key = _parcel.rec_name + _parcel.rec_tel
            pre_parcel = tmp_namelist.get(tmp_key)  # 根据当前的包裹的姓名从临时tmp_namelist字典里查找前一个同名的包裹
            if pre_parcel is None:  # 如果没有同名的包裹，则将当前的包裹放入临时的tmp_namelist字典
                tmp_namelist[tmp_key] = _parcel
            else:  # 有同名的key ，则对比两个包裹
                rep_id = select_repeat(pre_parcel, _parcel)
                if rep_id == 0:
                    print('无法确定')
                else:
                    # 如果重复件在tmp_namelist 里面，则用新的替换掉
                    if rep_id == pre_parcel.ser_id:
                        tmp_namelist.pop(tmp_key)
                        tmp_namelist[tmp_key] = _parcel
                    self.result_list.remove(rep_id)
                    self.repeat_list.append(rep_id)

    def filter_tel(self):
        tmp_tel_list = {}  # 临时电话字典 <tel,包裹>
        for cur_id in self.result_list:
            _parcel = self.dict_all.get(cur_id)
            cur_tel = _parcel.rec_tel
            pre_parcel = tmp_tel_list.get(cur_tel)
            if pre_parcel is None:
                tmp_tel_list[cur_tel] = _parcel
            else:
                rep_id = select_repeat(pre_parcel, _parcel)
                # 如果重复件在tmp_tel_list 里面，则用新的替换掉
                if rep_id == pre_parcel.ser_id:
                    tmp_tel_list.pop(cur_tel)
                    tmp_tel_list[cur_tel] = _parcel
                self.result_list.remove(rep_id)
                self.repeat_list.append(rep_id)

    def filter_address(self):
        tmp_address_list = {}  # 临时地址字典 <address,包裹>
        for cur_id in self.result_list:
            _parcel = self.dict_all.get(cur_id)
            cur_address = _parcel.rec_address
            pre_parcel = tmp_address_list.get(cur_address)
            if pre_parcel is None:
                tmp_address_list[cur_address] = _parcel
            else:
                rep_id = select_repeat(pre_parcel, _parcel)
                # 如果重复件在tmp_address_list 里面，则用新的替换掉
                if rep_id == pre_parcel.ser_id:
                    tmp_address_list.pop(cur_address)
                    tmp_address_list[cur_address] = _parcel
                self.result_list.remove(rep_id)
                self.repeat_list.append(rep_id)

    def set_title(self, sht, row, col, value):
        sht.Cells(row, col).Value = value
        sht.Cells(row, col).Font.Bold = True  # 是否黑体
        sht.Cells(row, col).Font.Size = 11  # 字体大小
        sht.Cells(row, col).Name = "宋体"  # 字体类型

    def set_cell(self, sht, row, col, value):
        sht.Cells(row, col).Value = value
        sht.Cells(row, col).Font.Size = 11  # 字体大小
        sht.Cells(row, col).Name = "宋体"  # 字体类型

    def save_file(self, save_path):
        curr_time = time.strftime("%Y%m%d%H%M%S", time.localtime())
        save_name = save_path + '发货数据%s.xls' % curr_time
        save_book = self.excel.Workbooks.Add()

        save_book.Worksheets.Add().Name = '重复件'
        rep_sheet = save_book.Worksheets('重复件')
        cnt = 1
        for no in self.repeat_list:
            parcel = self.dict_all.get(no)
            self.set_cell(rep_sheet, cnt, CONST.WPS_CUSTOMER_ROW, parcel.ser_id)
            self.set_cell(rep_sheet, cnt, CONST.WPS_CUSTOMER_ROW+1, parcel.customer)
            self.set_cell(rep_sheet, cnt, CONST.WPS_CUSTOMER_CODE_ROW+1, parcel.customer_code)
            self.set_cell(rep_sheet, cnt, CONST.WPS_DATE_ROW+1, parcel.date)
            self.set_cell(rep_sheet, cnt, CONST.WPS_NUM_ROW+1, parcel.num)
            self.set_cell(rep_sheet, cnt, CONST.WPS_TRANS_CODE_ROW+1, parcel.num)
            self.set_cell(rep_sheet, cnt, CONST.WPS_NAME_ROW+1, parcel.rec_name)
            self.set_cell(rep_sheet, cnt, CONST.WPS_REV_PRINCE_ROW+1, parcel.rec_privince)
            self.set_cell(rep_sheet, cnt, CONST.WPS_REV_CITY_ROW+1, parcel.rec_city)
            self.set_cell(rep_sheet, cnt, CONST.WPS_ADDRESS_ROW+1, parcel.rec_address)
            self.set_cell(rep_sheet, cnt, CONST.WPS_ADDRESS_CODE_ROW+1, parcel.address_code.zfill(6))
            self.set_cell(rep_sheet, cnt, CONST.WPS_TEL_ROW+1, parcel.rec_tel)
            self.set_cell(rep_sheet, cnt, CONST.WPS_WGT_ROW+1, parcel.weight)
            self.set_cell(rep_sheet, cnt, CONST.WPS_GOODS_ROW+1, parcel.goods)
            self.set_cell(rep_sheet, cnt, CONST.WPS_DESTINATION_ROW+1, parcel.distination)
            self.set_cell(rep_sheet, cnt, CONST.WPS_SENDER_NAME_ROW+1, parcel.sender)
            self.set_cell(rep_sheet, cnt, CONST.WPS_SENDER_TEL_ROW+1, parcel.sender_tel)
            self.set_cell(rep_sheet, cnt, CONST.WPS_REV_EMAIL_ROW+1, parcel.rev_email)
            self.set_cell(rep_sheet, cnt, CONST.WPS_STATE_ROW+1, parcel.state)
            cnt += 1

        save_book.Worksheets.Add().Name = 'LIST'
        sheet = save_book.Worksheets('LIST')
        # 设置TITLE
        col = 2
        for title in CONST.TITLE_LIST:
            self.set_title(sheet, 1, col, title)
            col += 1

        # 设置内容
        row = 2
        for no in self.result_list:
            parcel = self.dict_all.get(no)
            self.set_cell(sheet, row, CONST.WPS_CUSTOMER_ROW, parcel.ser_id)
            self.set_cell(sheet, row, CONST.WPS_CUSTOMER_ROW+1, parcel.customer)
            self.set_cell(sheet, row, CONST.WPS_CUSTOMER_CODE_ROW+1, parcel.customer_code)
            self.set_cell(sheet, row, CONST.WPS_DATE_ROW+1, parcel.date)
            self.set_cell(sheet, row, CONST.WPS_NUM_ROW+1, parcel.num)
            self.set_cell(sheet, row, CONST.WPS_TRANS_CODE_ROW+1, parcel.num)
            self.set_cell(sheet, row, CONST.WPS_NAME_ROW+1, parcel.rec_name)
            self.set_cell(sheet, row, CONST.WPS_REV_PRINCE_ROW+1, parcel.rec_privince)
            self.set_cell(sheet, row, CONST.WPS_REV_CITY_ROW+1, parcel.rec_city)
            self.set_cell(sheet, row, CONST.WPS_ADDRESS_ROW+1, parcel.rec_address)
            self.set_cell(sheet, row, CONST.WPS_ADDRESS_CODE_ROW+1, parcel.address_code)
            self.set_cell(sheet, row, CONST.WPS_TEL_ROW+1, parcel.rec_tel)
            self.set_cell(sheet, row, CONST.WPS_WGT_ROW+1, parcel.weight)
            self.set_cell(sheet, row, CONST.WPS_GOODS_ROW+1, parcel.goods)
            self.set_cell(sheet, row, CONST.WPS_DESTINATION_ROW+1, parcel.distination)
            self.set_cell(sheet, row, CONST.WPS_SENDER_NAME_ROW+1, parcel.sender)
            self.set_cell(sheet, row, CONST.WPS_SENDER_TEL_ROW+1, parcel.sender_tel)
            self.set_cell(sheet, row, CONST.WPS_REV_EMAIL_ROW+1, parcel.rev_email)
            self.set_cell(sheet, row, CONST.WPS_STATE_ROW+1, parcel.state)
            row += 1

        save_book.SaveAs(save_name)


path = r'C:\Users\Hunter\PycharmProjects\starex\\'
file_name = '快件核对_0507112211.xls'

wps = WPS(path + file_name)
parcel_list = wps.read_xls()
wps.filter_name()
wps.filter_tel()
wps.filter_address()

wps.result_list.sort()
wps.repeat_list.sort()

wps.save_file(path)

print(wps.result_list.__len__(),'self.result_list', wps.result_list)
print(wps.repeat_list.__len__(),'self.repeat_list', wps.repeat_list)





# filter_parcel()
# for _parcel in parcel_list:
#     print(':', _parcel.num)
# xlApp = Dispatch('Excel.Application')
# xlApp.Visible = True
# xlApp.Workbooks.Add()
# xlApp.Worksheets.Add().Name = 'test'
# xlSheet = xlApp.Worksheets('test')
# xlSheet.Cells(1,1).Value = 'title'
# xlSheet.Cells(2,1).Value = 123