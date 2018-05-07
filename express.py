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

    def set_customer(self, customer):
        self.customer = customer

    def set_customer_code(self, customer_code):
        self.customer_code = customer_code

    def set_rec_privince(self, rec_privince):
        self.rec_privince = rec_privince

    def set_rec_city(self, rec_city):
        self.rec_city = rec_city

    def set_address_code(self, address_code):
        self.address_code = address_code

    def set_address_code(self, address_code):
        self.address_code = address_code

    def set_distination(self, distination):
        self.distination = distination

    def set_sender(self, sender):
        self.sender = sender

    def set_sender_tel(self, sender_tel):
        self.sender_tel = sender_tel

    def set_rev_email(self, rev_email):
        self.rev_email = rev_email

    def set_remark1(self, remark1):
        self.remark1 = remark1

    def set_remark2(self, remark2):
        self.remark2 = remark2

    def set_state(self, state):
        self.state = state

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
