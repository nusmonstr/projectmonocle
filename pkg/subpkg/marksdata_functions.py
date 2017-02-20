"""
Created on Aug 23, 2016

@author: Eric

Functions:
import_citi
import_fidelity
import_healthequity
import_vanguard
import_adp
import_widget
_widgetwhat
_categorizeADP
_categorizeFidelity
_categorizeHealthequity
_categorizeVanguard

"""
from csv import reader
from re import findall
from itertools import islice
from .financialelements import Bank, Transaction
import os
from datetime import timedelta, datetime


def import_csv(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, skip_header, None):
            # Verify that the line is not empty before unpacking
            if line:
                state, pubdate, desc, deb, cred, *trash = line
                amnt = '-' + deb if deb is not '' else cred
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, acnt=bank_name, status=bank_status)
                s.cat = ''
                s.subcat = ''
                s.notes = ''
                s.tag = ''
                statement.append(s)
    return statement, values


def import_umcu(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, skip_header, None):
            # Verify that the line is not empty before unpacking
            if line:
                _, pubdate, amnt, _, cat, desc, _, notes = line
                #amnt = '-' + deb if deb is not '' else cred
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, acnt=bank_name, status=bank_status)
                s.cat = cat
                s.subcat = ''
                s.notes = notes
                s.tag = ''
                statement.append(s)
    return statement, values


def import_capitalone(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, skip_header, None):
            # Verify that the line is not empty before unpacking
            if line:
                _, _, pubdate, _, desc, cat, deb, cred = line
                amnt = '-' + deb if deb is not '' else cred
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, acnt=bank_name, status=bank_status)
                s.cat = cat
                s.subcat = ''
                s.notes = ''
                s.tag = ''
                statement.append(s)
    return statement, values


def import_qif(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    fh = open(bank_data, 'r')
    lines = fh.readlines()
    # $$$ Make this into a function for easier troubleshooting - Convert qif formatting to csv formatting
    lines = ''.join(lines).replace('\nT', '\t').replace('\nM', ',').replace('\nD', ',').replace('\n^', '').replace('\n', '').split('\t')
    for line in lines[skip_header:]:
        # Verify that the line is not empty before unpacking
        if [line]:
            amnt, pubdate, desc, *trash = line.split(',')
            # Create new Transaction object with formatted data
            s = Transaction(pubdate, desc, amnt, acnt=bank_name, status=bank_status)
            s.cat, s.subcat = _categorize_by_description(s.desc)
            s.notes = ''
            s.tag = ''
            statement.append(s)
    return statement, values


def import_qfx(bank):
    # Need to pull both checking Transactions then repeat for savings
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    keys = ['DTPOSTED', 'MEMO', 'TRNAMT', 'NAME']
    fh = open(bank_data, 'r')
    lines = fh.readlines()
    lines = ''.join(lines).replace('\n', '')
    line_Val = findall(r'<LEDGERBAL>(.*?)</LEDGERBAL>', lines)
    lines = findall(r'<STMTTRN>(.*?)</STMTTRN>', lines)
    for line in lines:
        line = line[1:].replace('<', ',').replace('>', ',').split(',')
        tags, target = line[::2], line[1::2]  # Return tags and values from an alternating sequence
        pairs = dict(zip(tags, target))
        pubdate, desc, amnt, name = [pairs.get(key) for key in keys]
        pubdate = pubdate + 'Ymd'
        # Create new Transaction object with formatted data
        s = Transaction(pubdate, desc, amnt, acnt=bank_name, status=bank_status)
        s.cat = ''
        s.subcat = ''
        s.notes = ''
        s.tag = ''
        statement.append(s)
    # Import current balance information
    balance_date = datetime.date(datetime.fromtimestamp(os.path.getmtime(bank_data)))
    balance = line_Val[0].replace('<', ',').replace('>', ',').split(',')
    v = Transaction(balance_date, 'Placeholder', balance[2], acnt=bank_name, status=bank_status)
    values.append(v)
    return statement, values


def _categorize_by_description(what):
    cat = ''
    subcat = ''
    if what.startswith('Contribu') : cat,subcat = 'Transfer','Account'
    elif what.startswith('Dividend') : cat,subcat = 'Income','Dividends'
    return [cat, subcat]
