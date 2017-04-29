#!/usr/local/bin/python3

import imaplib, getpass, email, pprint, argparse, re, sys, ssl

from email import policy as pl
from dateutil.parser import parse as date_parse
from decimal import Decimal
from openpyxl import Workbook


def process_mailbox(mb_reader, s=r"(SINCE 1-Mar-2017 FROM 900)"):
    """
    Search messages from Sberbank in mailbox    
    :param mb_reader: mailbox reader object
    :param s: search string in IMAP format
    :return: messages as list of dicts {'time':headers, 'body':body}, time is datetime object, body is a string
    """

    def stop_words(message):
        for word in ('пароль', 'вход в сбербанк'):
            if word in message['body'].lower():
                return False
        return True

    sms_list = []

    print("Searching...")
    rv, data = mb_reader.search(None, s)
    if rv != 'OK':
        print("No messages found!")
        return sms_list

    mset = ""
    for d in data[0].split():
        mset = mset + "," + d.decode('ascii')

    print("Fetching...")
    rv, data = mb_reader.fetch(mset, '(RFC822)')
    if rv != 'OK':
        print("ERROR getting message", mset)
        return sms_list

    msg_list = []
    for m in data:
        if len(m) > 1:
            msg_list.append(m[1])

    print("Processing...")
    counter = 0
    for msg_bytes in msg_list:
        msg = email.message_from_bytes(msg_bytes, policy=pl.default)
        for part in msg.walk(): # Parse message
            sms_list.append({'time': date_parse(dict(part.items())['Date']), 'body': part.get_content()})
        counter += 1
        if counter % 25 == 0: # Make some awaiting progress
            print("Processed ", counter, "messages")

    return list(filter(stop_words, sms_list))


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

    for transaction in trans_list:
        try:
            values = re.match(
                r'(.+?) ([0-9]+\.[0-9]+\.[0-9]+ [0-9]+:[0-9]+) (.+?) ([0-9]+(?:\.[0-9]+)*)(.+?)(?: с комиссией ([0-9]+(?:\.[0-9]+)*)(.+?))?( .+)? Баланс: ([0-9]+(?:\.[0-9]+)*)(?:.+)',
                transaction['body'])
            if values: # Purchases, ATM operations and another incomes&expences
                oper.append({
                    'card': values.group(1),
                    'time1': date_parse(values.group(2), dayfirst=True),
                    'oper': values.group(3),
                    'sum': Decimal(values.group(4)),
                    'currency': values.group(5),
                    'comission': Decimal(values.group(6)) if values.group(6) else None,
                    'commcurr': values.group(7),
                    'place': values.group(8),
                    'bal': Decimal(values.group(9)),
                    'time': transaction['time']
                })
                continue
            values = re.match(
                r'(.+?) ([0-9]+\.[0-9]+\.[0-9]+) (.+) ([0-9]+(?:\.[0-9]+)*)(.+?) Баланс: ([0-9]+(?:\.[0-9]+)*)(?:.+)',
                transaction['body']
            )
            if values:
                oper.append({
                    'card': values.group(1),
                    'time1': date_parse(values.group(2), dayfirst=True),
                    'oper': values.group(3),
                    'sum': Decimal(values.group(4)),
                    'currency': values.group(5),
                    'comission': None,
                    'commcurr': None,
                    'place': None,
                    'bal': Decimal(values.group(6)),
                    'time': transaction['time']
                })
                continue
            values = re.match(
                r'Сбербанк Онлайн. (.+?) перевел(?:.+?) ([0-9]+(?:\.[0-9]+)*) ([^ .]+)\.?(?: Сообщение: "?([^"]+)"?)?',
                transaction['body']
            )
            if values:
                trf.append({
                    'name': values.group(1),
                    'sum': Decimal(values.group(2)),
                    'currency': values.group(3),
                    'comment': values.group(4),
                    'time': transaction['time']
                })
                continue
            values = re.match(
                r'(.+?) ([0-9.:]+) (.+) ([0-9]+(?:\.[0-9]+)*)(.+?)\.? от отправителя (.+)',
                transaction['body']
            )
            if values:
                trf.append({
                    'name': values.group(6),
                    'sum': Decimal(values.group(4)),
                    'currency': values.group(5),
                    'comment': "",
                    'time': transaction['time']
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


def process_arguments():
    """
    Processes command line arguments 
    :return: dict of configuration options
    """

    parser = argparse.ArgumentParser(description="Process Sberbank SMS messages backed up to imap server and generate"
                                                 "xlsx sheet",
                                     formatter_class=argparse.ArgumentDefaultsHelpFormatter
                                     )
    parser.add_argument("-d", "--date", help="Start from date", default="1-Mar-2017")
    parser.add_argument("-l", "--login", help="Login with this name")
    parser.add_argument("-p", "--password", help="Login with this password")
    parser.add_argument("-s", "--site", help="Connect to this imap server", default="imap.gmail.com")
    parser.add_argument("-f", "--folder", help="Folder to read SMS from", default="SMS")
    parser.add_argument("-S", "--search", help="IMAP search string", default="FROM 900")
    parser.add_argument("-w", "--warn", help="print warnings", action="store_true")
    parser.add_argument("outfile", help="Output MS Excel file, please add .xlsx explicitly")

    return vars(parser.parse_args())


def save_operations(arg, wb_file='sbercards.xlsx'):
    """
    Save operations to xlsx
    :param wb_file: write to this file
    :param arg: tuple of (oper, trf), oper is list of transactions as sms-es
            trf is list of transfers as sms-es
    :return: None
    """

    oper, trf = arg

    wb = Workbook()
    ws = wb.active
    ws.title = "Operations"

    i = 1
    for val in ("Card", "Time", "Time inside SMS", "Operation", "Sum", "Currency",
                "Comission", "Comm. currency", "Balance", "Place",
                "Name", "Comment", "Time of transfer"
               ):
        ws.cell(row=1, column=i, value=val)
        i += 1

    cur_row = 2
    for o in oper:
        ws.cell(row=cur_row, column=1, value=o['card'])
        ws.cell(row=cur_row, column=2, value=o['time'])
        ws.cell(row=cur_row, column=3, value=o['time1'])
        ws.cell(row=cur_row, column=4, value=o['oper'])
        ws.cell(row=cur_row, column=5, value=o['sum'])
        ws.cell(row=cur_row, column=6, value=o['currency'])
        ws.cell(row=cur_row, column=7, value=o['comission'])
        ws.cell(row=cur_row, column=8, value=o['commcurr'])
        ws.cell(row=cur_row, column=9, value=o['bal'])
        ws.cell(row=cur_row, column=10, value=o['place'])
        if ('transfer' not in o.keys()) or ('transfer' in o.keys() and o['transfer'] == 'Not found'):
            ws.cell(row=cur_row, column=11, value="")
            ws.cell(row=cur_row, column=12, value="")
            ws.cell(row=cur_row, column=13, value="")
        else:
            ws.cell(row=cur_row, column=11, value=o['transfer']['name'])
            ws.cell(row=cur_row, column=12, value=o['transfer']['comment'])
            ws.cell(row=cur_row, column=13, value=o['transfer']['time'])
        cur_row += 1

    ws1 = wb.create_sheet("Transfers")

    i = 1
    for val in ('Time', 'Name', 'Sum', 'Comment'):
        ws1.cell(row=1, column=i, value=val)
        i += 1

    cur_row = 2
    for transaction in trf:
        ws1.cell(row=cur_row, column=1, value=transaction['time'])
        ws1.cell(row=cur_row, column=2, value=transaction['name'])
        ws1.cell(row=cur_row, column=3, value=transaction['sum'])
        ws1.cell(row=cur_row, column=4, value=transaction['comment'])
        cur_row += 1

    wb.save(wb_file)

if __name__ == "__main__":

    config_opts = process_arguments()

    context = ssl.create_default_context()
    context.check_hostname = False
    context.verify_mode = ssl.CERT_NONE
    mb_reader = imaplib.IMAP4_SSL(config_opts['site'], ssl_context=context)

    try:
        rv, data = mb_reader.login(config_opts['login'],
                                   config_opts['password'] if config_opts['password'] else getpass.getpass()
                                   )
    except imaplib.IMAP4.error:
        print("LOGIN FAILED!!! ")
        sys.exit(1)

    rv, data = mb_reader.select(config_opts['folder'])
    if rv == 'OK':
        search_string = "(" + config_opts['search'] + " SINCE " + config_opts['date'] + ")"
        sms_list = process_mailbox(mb_reader, search_string)
        mb_reader.close()
    else:
        print("ERROR: Unable to open mailbox ", rv)
        sys.exit(1)

    if sms_list:
        save_operations((process_sms_list(sms_list, warn=config_opts['warn'])), wb_file=config_opts['outfile'])

    mb_reader.logout()
