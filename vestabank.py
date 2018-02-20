#!/usr/local/bin/python3

import re, pprint
from decimal import Decimal
from dateutil.parser import parse as date_parse


def stop_words(message):
    for word in ('карта', 'karta'):
        if word in message['body'].lower():
            for word in ('otrazhena v vypiske', 'vhod v internet-bank', 'вход в vestabank', \
                         'вход в мобильное приложение', 'ispolnen platezh', 'пароль',):
                if word in message['body'].lower():
                    return False
            return True

    return False

def process_sms_list(trans_list, warn=False):
    """
    Make list of operations from list of SMS
    :param warn: print warnings
    :param trans_list: list of transactions as dicts with 'time' and 'body' keys
    :return: tuple of (operations, transfers), operations as list of dicts with 'card', 'time', time1', 'time2', 'operation', 'currency',
            'sum', 'balance', 'person', 'place', 'name', 'sum1', 'comment' keys
            transfers as list of dicts with 'name', 'sum', 'comment', 'time' keys
    """

    oper = []  # Card operations
    trf = []  # Money transfers

    purchase_re = re.compile(r'^(?:Karta|Карта) ([0-9]+?): (.+?), (.+?) ([0-9.]+) (.+?)[.,] (?:(?:комиссия|komissiya) D([0-9.]+) (.+?)\. )?(?:(.+?)\. )? *(?:Доступно|Dostupno) ([0-9.]+) (.+?)\.')

    for transaction in trans_list:
        try:
            values = purchase_re.match(transaction['body'])
            if values: # Purchases, ATM operations and another incomes&expences
                d = {
                    'card': values.group(1),
                    'time1': date_parse(values.group(2), dayfirst=True),
                    'oper': values.group(3),
                    'sum': Decimal(values.group(4)) if values.group(4) else None,
                    'currency': values.group(5),
                    'comission': Decimal(values.group(6)) if values.group(6) else None,
                    'commcurr': values.group(7),
                    'place': values.group(8),
                    'bal': Decimal(values.group(9)) if values.group(9) else None,
                    'time': transaction['time']
                }
                oper.append(d)
                continue
            if warn:
                print("WARNING: unknown transaction")
                pprint.pprint(transaction)
        except:
            print("ERROR: unable to process")
            pprint.pprint(transaction)

    return oper, trf

if __name__ == "__main__":
    print("This module is for import only")
