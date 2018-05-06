#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__author__ = 'Hunter'

class Parcel(object):
    def __init__(self, ser_id, num, rec_name, rec_tel, rec_address):
        self.ser_id = ser_id
        self.num = num
        self.rec_name = rec_name
        self.rec_tel = rec_tel

        if rec_address is not None:
            self.rec_address = rec_address.replace('\u2212', '-')

    def set_date(self, date):
        self.date = date

    def set_goods(self, goods):
        self.goods = goods

    def set_weight(self, weight):
        self.weight = weight

    def print_parcel(self):
        print(self.num)



if __name__ == '__main__':
    parcel = Parcel(1, '121343', 'ag', '13', '安静')
    parcel.print_parcel()
