#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__author__ = 'Hunter'

class Parcel(object):
    def __init__(self, num, rec_name, rec_tel, rec_address):
        self.num = num
        self.rec_name = rec_name
        self.rec_tel = rec_tel
        if rec_address is not None:
            self.rec_address = rec_address.replace('\u2212', '-')

    def print_parcel(self):
        print(self.num)



if __name__ == '__main__':
    parcel = Parcel('121343', 'ag', '13', '安静')
    parcel.print_parcel()
