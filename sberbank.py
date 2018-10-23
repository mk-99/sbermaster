#!/usr/local/bin/python3

import re, pprint
from decimal import Decimal
from dateutil.parser import parse as date_parse


def stop_words(message):
    for word in ('пароль', 'вход в сбербанк'):
        if word in message['body'].lower():
            return False
    return True


def process_sms_list(trans_list, warn=False):
    """
    Make list of operations from list of SMS
    :param warn: print warnings
    :param trans_list: list of transactions as dicts with 'time' and 'body' keys
    :return: tuple of (operations, transfers), operations as list of dicts with 'card', 'time', time1', 'time2', 'operation', 'currency',
            'sum', 'balance', 'person', 'place', 'name', 'sum1', 'comment' keys
            transfers as list of dicts with 'name', 'sum', 'comment', 'time' keys
    """
    def find_transfer(o, trf):
        """
        Finds transfer in trf that best matches operation o. "Best" means minimal time shift and sum equality
        :param o: operation
        :param trf: list of transfers
        :return: transfer the most relevant to operation
        """
        MAXDELTA = 150 # maximum seconds between transaction and operation SMS-es
        candidates = []
        min_seconds = 'Unknown'
        retval = 'Not found'

        for t in trf:
            if t['sum'] == o['sum']:
                candidates.append(t)
        if not len(candidates):
            return retval

        for c in candidates:
            delta = abs(int((c['time'] - o['time']).total_seconds()))
            if delta > MAXDELTA:
                continue
            if min_seconds == 'Unknown':
                min_seconds = delta
                retval = c
            elif min_seconds > delta:
                min_seconds = delta
                retval = c
            else:
                retval = c

        return retval

    oper = []  # Card operations
    trf = []  # Money transfers

    purchase_re = re.compile(r'(.+?) ((?:[0-9]+\.[0-9]+\.[0-9]+ )?[0-9]+:[0-9]+) (.+?) ([0-9]+(?:\.[0-9]+)*)(.+?)(?: с комиссией ([0-9]+(?:\.[0-9]+)*)(.+?))?( .+)? Баланс: ([0-9]+(?:\.[0-9]+)*)(?:.+)')
    mobilebank_re = re.compile(r'(.+?) ([0-9]+\.[0-9]+\.[0-9]+) (.+) ([0-9]+(?:\.[0-9]+)*)(.+?) Баланс: ([0-9]+(?:\.[0-9]+)*)(?:.+)')
    transfer_re = re.compile(r'Сбербанк Онлайн. (.+?) перевел(?:.+?) ([0-9]+(?:\.[0-9]+)*) ([^ .]+)\.?(?: Сообщение: "?([^"]+)"?)?')
    receive_re = re.compile(r'(.+?):? ([0-9.:]+) (.+) ([0-9]+(?:\.[0-9]+)*)(.+?)\.? от отправителя (.+)(?: Сообщение: "?([^"]+)"?)')
    receive2_re= re.compile(r'(.+?) ([0-9.:]+) (.+) ([0-9]+(?:\.[0-9]+)*)(.+?)\.? от отправителя (.+)')

    for transaction in trans_list:
        try:
            values = purchase_re.match(transaction['body'])
            if values: # Purchases, ATM operations and another incomes&expences
                oper.append({
                    'time': transaction['time'],
                    'card': values.group(1),
                    'time1': date_parse(values.group(2), default=transaction['time'], dayfirst=True),
                    'oper': values.group(3),
                    'sum': Decimal(values.group(4)),
                    'currency': values.group(5),
                    'comission': Decimal(values.group(6)) if values.group(6) else None,
                    'commcurr': values.group(7),
                    'place': values.group(8),
                    'bal': Decimal(values.group(9))
                })
                continue
            values = mobilebank_re.match(transaction['body'])
            if values:
                oper.append({
                    'time': transaction['time'],
                    'card': values.group(1),
                    'time1': date_parse(values.group(2), default=transaction['time'], dayfirst=True),
                    'oper': values.group(3),
                    'sum': Decimal(values.group(4)),
                    'currency': values.group(5),
                    'comission': None,
                    'commcurr': None,
                    'place': None,
                    'bal': Decimal(values.group(6))
                })
                continue
            values = transfer_re.match(transaction['body'])
            if values:
                trf.append({
                    'time': transaction['time'],
                    'name': values.group(1),
                    'sum': Decimal(values.group(2)),
                    'currency': values.group(3),
                    'comment': values.group(4)
                })
                continue
            values = receive_re.match(transaction['body'])
            if values:
                trf.append({
                    'time': transaction['time'],
                    'name': values.group(6),
                    'sum': Decimal(values.group(4)),
                    'currency': values.group(5),
                    'comment': values.group(7)
                })
                continue
            values = receive2_re.match(transaction['body'])
            if values:
                trf.append({
                    'time': transaction['time'],
                    'name': values.group(6),
                    'sum': Decimal(values.group(4)),
                    'currency': values.group(5),
                    'comment': ""
                })
                continue
            if warn:
                print("WARNING: unknown transaction")
                pprint.pprint(transaction)
        except:
            print("ERROR: unable to process")
            pprint.pprint(transaction)

    for o in oper:
        if o['oper'] == 'зачисление':
            o['transfer'] = find_transfer(o, trf)

    return oper, trf

if __name__ == "__main__":
    print("This module is for import only")
