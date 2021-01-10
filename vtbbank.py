#!/usr/local/bin/python3

import re, pprint
from decimal import Decimal
from dateutil.parser import parse as date_parse


def stop_words(message):
    for word in ('карта', 'karta'):
        if word in message['body'].lower():
            for word in ('nikomu ne', 'vhod v', \
                         'вход в', 'пароль',):
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

# Karta *8741: Oplata 250.00 RUB;IP SOROKIN E.A. SMT;21.10.2018 17:05,dostupno 283.16 RUB

    purchase_re = re.compile(r'^Karta \*([0-9]+?): (.+?) ([0-9.]+) (.+?);(.+?);(.+)[,;] ?dostupno ([0-9.]+) ([^.]+)(?:\(.+\))?\.?$')
    refund_re = re.compile(r'^Karta \*([0-9]+?): (.+?) ([0-9.]+) (.+?); ?dostupno ([0-9.]+) ([^.]+).+$')

    purchase2_re = re.compile(r'^(.+?) ([0-9.]+)(.+?) (?:Karta|Карта)\*(.+?) (.+?) (?:Balans|Баланс) ([0-9.]+)(.+?) ([0-9]+:[0-9]+)')

    for transaction in trans_list:
        try:
            values = purchase_re.match(transaction['body'])
            if values: # Purchases, ATM operations and another incomes&expences
                d = {
                    'time': transaction['time'],
                    'card': values.group(1),
                    'time1': date_parse(values.group(6), dayfirst=True),
                    'oper': values.group(2),
                    'sum': Decimal(values.group(3)) if values.group(3) else None,
                    'currency': values.group(4),
                    'comission': None,
                    'commcurr': None,
                    'place': values.group(5).strip(),
                    'bal': Decimal(values.group(7)) if values.group(7) else None
                }
                oper.append(d)
                continue
            values = refund_re.match(transaction['body'])
            if values: # Refunds and another deposits
                d = {
                    'time': transaction['time'],
                    'card': values.group(1),
                    'time1': None,
                    'oper': values.group(2),
                    'sum': Decimal(values.group(3)) if values.group(3) else None,
                    'currency': values.group(4),
                    'comission': None,
                    'commcurr': None,
                    'place': None,
                    'bal': Decimal(values.group(5)) if values.group(5) else None
                }
                oper.append(d)
                continue
            values = purchase2_re.match(transaction['body'])
            if values: # Refunds and another deposits
                d = {
                    'time': transaction['time'],
                    'card': values.group(4),
                    'time1': date_parse(values.group(8), default=transaction['time'], dayfirst=True),
                    'oper': values.group(1),
                    'sum': Decimal(values.group(2)) if values.group(2) else None,
                    'currency': values.group(3),
                    'comission': None,
                    'commcurr': None,
                    'place': values.group(5),
                    'bal': Decimal(values.group(6)) if values.group(6) else None
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
