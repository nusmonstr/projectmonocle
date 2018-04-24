"""
Created on Aug 23, 2016

@author: Eric Champe

Functions:
import_citi
import_umb_bank
import_fidelity
import_vanguard
import_widget
_widgetwhat
_categorizeFidelity
_categorizeVanguard

"""
from csv import reader
from re import findall
from itertools import islice
from .financialelements import Bank, Transaction
import os
from datetime import timedelta, datetime, date
import win32com.client

userpayp = 'Semimonthly'

vanguard_fund_lookup_1 = {
    'CASH': 'External Account',
    'VANGUARD 500 INDEX ADMIRAL CL': '500 Index Fund Adm',
    'VANGUARD FEDERAL MONEY MARKET FUND': 'Federal Money Market',
    'VANGUARD HEALTHCARE INVESTOR CL': 'Healthcare Index Fund',
    'VANGUARD LONG TERM GOVT BOND INDEX ADMIRAL CL': 'Longterm Gov Bond Index',
    'VANGUARD MID CAP INDEX ADMIRAL CL': 'Mid-Cap Index Fund Adm',
    'VANGUARD PRECIOUS METALS & MINING INVESTOR CL': 'Precious Metals & Mining',
    'VANGUARD SMALL CAP INDEX ADMIRAL CL': 'Small-Cap Index Fund Adm',
    'VANGUARD TOTAL BOND MARKET INDEX ADMIRAL CL': 'Total Bond Mkt Index Adm',
    'VANGUARD TOTAL INTL STOCK INDEX ADMIRAL CL': 'Tot Intl Stock Ix Admiral',
    'VANGUARD WELLINGTON INVESTOR CL': 'Wellington Fund Inv',
    'VANGUARD WELLINGTON ADMIRAL CL' : 'Wellington Fund Adm',
    'VANGUARD LONG TERM TREASURY INDEX ADMIRAL CL': 'Longterm Gov Bond Index'}

vanguard_fund_lookup_2 = {
    'Vanguard 500 Index Fund Admiral Shares': '500 Index Fund Adm',
    'Vanguard Federal Money Market Fund': 'Federal Money Market',
    'Vanguard Health Care Fund Investor Shares': 'Healthcare Index Fund',
    'Vanguard Long-Term Government Bond Index Fund Admiral Shares': 'Longterm Gov Bond Index',
    'Vanguard Mid-Cap Index Fund Admiral Shares': 'Mid-Cap Index Fund Adm',
    'Vanguard Precious Metals And Mining Fund Investor Shares': 'Precious Metals & Mining',
    'Vanguard Precious Metals and Mining Fund': 'Precious Metals & Mining',
    'Vanguard Small-Cap Index Fund Admiral Shares': 'Small-Cap Index Fund Adm',
    'Vanguard Total Bond Market Index Fund Admiral Shares': 'Total Bond Mkt Index Adm',
    'Vanguard Total International Stock Index Fund Admiral Shares': 'Tot Intl Stock Ix Admiral',
    'Vanguard Wellington Fund Investor Shares': 'Wellington Fund Inv',
    'Vanguard Wellington\x1a Fund Investor Shares': 'Wellington Fund Inv',
    'Vanguard Wellington\x1a Fund Admiral\x1a Shares':  'Wellington Fund Adm',
    'Vanguard Wellington Fund Admiral Shares':  'Wellington Fund Adm',
    'Vanguard Long-Term Treasury Index Fund Admiral Shares': 'Longterm Gov Bond Index'}


def import_citi(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, skip_header, None):
            if line:
                state, pubdate, desc, deb, cred, *trash = line
                amnt = '-' + deb if deb is not '' else cred
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
                s.cat = s.sub_cat = ''
                s.notes = ''
                s.tag = ''
                statement.append(s)
    return statement, list()


def import_umbbank(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, skip_header, None):
            if line:
                pubdate, trans_type, desc, memo, amnt = line
                desc = desc + memo.strip()
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
                s.cat = s.sub_cat = _categorizeUMB(desc)
                s.notes = ''
                s.tag = ''
                statement.append(s)
    return statement, list()


def import_fidelity(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, skip_header, None):
            if line:
                pubdate, fund, desc, amnt, shares, *trash = line
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
                s.cat, s.sub_cat = _categorizeFidelity(s.desc)
                s.notes = 'Shares/Unit '+shares+' | '+fund
                s.tag = ''
                statement.append(s)
    return statement, list()


def import_firsttech(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    balance_values = list()
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, skip_header, None):
            if line:
                id, pubdate, effdate, type, amnt, check_no, ref, desc, tran_cat, method, balance = line
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
                s.cat = s.sub_cat = ''
                s.notes = ''
                s.tag = ''
                statement.append(s)
                balance_values.append((pubdate, balance))
    baldate = datetime.date(datetime.fromtimestamp(os.path.getmtime(bank_data)))
    balance = round(float(sorted(balance_values, key=lambda x: x[0], reverse=True)[0][1]), 2)
    v = Transaction(baldate, 'ACCOUNT BALANCE GAP', balance, paypstyle=userpayp, acnt=bank_name, status=bank_status)
    return statement, [v]


def import_widget(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    keys = ['DTPOSTED', 'MEMO', 'TRNAMT', 'NAME']
    fh = open(bank_data, 'r')
    lines = fh.readlines()
    lines = ''.join(lines).replace('\n', '')
    # for each in lines: print(each)
    line_Val = findall(r'<LEDGERBAL><BALAMT>(.*?)</LEDGERBAL>', lines)
    lines = findall(r'<STMTTRN>(.*?)</STMTTRN>', lines)
    for line in lines:
        line = line[1:].replace('<', ',').replace('>', ',').split(',')
        tags, target = line[::2], line[1::2]  # Return tags and values from an alternating sequence
        pairs = dict(zip(tags, target))
        pubdate, desc, amnt, name = [pairs.get(key) for key in keys]
        pubdate = pubdate + 'Ymd'
        desc = _widgetwhat(desc, name)
        # Create new Transaction object with formatted data
        s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
        s.cat = s.sub_cat = ''
        s.notes = ''
        s.tag = ''
        statement.append(s)
    baldate = datetime.date(datetime.fromtimestamp(os.path.getmtime(bank_data)))
    balance = line_Val[0].replace('<', ',').split(',')[0]
    v = Transaction(baldate, 'ACCOUNT BALANCE GAP', balance, paypstyle=userpayp, acnt=bank_name, status=bank_status)
    values.append(v)
    return statement, values


def import_vanguard(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    with open(bank_data) as f:
        contents = [x.strip().replace(',', '') for x in f.readlines()]
        val_start = contents.index('Account Number,Investment Name,Symbol,Shares,Share Price,Total Value,'.replace(',', '')) + 1
        trans_start = contents.index('Account Number,Trade Date,Settlement Date,Transaction Type,Transaction Description,Investment Name,Symbol,Shares,Share Price,Principal Amount,Commission Fees,Net Amount,Accrued Interest,Account Type,'.replace(',', '')) + 1
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, trans_start, None):
            if ''.join(line):
                account, when0, pubdate, trantype, desc, fund, symbol, shares, price, amnt, *trash = line
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
                if account == '73510370':
                    s.acnt = 'Vanguard'
                elif account == '46069832':
                    s.acnt = 'Vanguard - Roth Brokerage'
                s.cat, s.sub_cat = _categorizeVanguard(desc)
                s.notes = 'Share Price: '+price+' Shares: '+shares
                s.fund = vanguard_fund_lookup_1[fund]
                if s.pubdate > date(year=2018, month=2, day=1):
                    statement.append(s)
    pubdate = datetime.date(datetime.fromtimestamp(os.path.getmtime(bank_data)))
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, val_start, trans_start-1):
            # Verify that the line is not empty before unpacking
            if ''.join(line):
                account, desc, symbol, shares, price, amnt, *trash = line
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
                if account == '73510370':
                    s.acnt = 'Vanguard'
                elif account == '46069832':
                    s.acnt = 'Vanguard - Roth Brokerage'
                else:
                    continue    # Don't include transactions from defunct Accounts
                s.cat, s.sub_cat = ('Income', 'Valuation')
                s.notes = 'Share Price: '+price+' Shares: '+shares
                s.fund = vanguard_fund_lookup_2[desc]
                values.append(s)
    return statement, values


def _widgetwhat(what, name):
    if what == '': what = name if name != '' else 'Description not Found'
    elif what == 'Deposit' or what == 'Withdrawal' : pass
    elif what.startswith('Deposit at '): pass
    elif what.startswith('Deposit by '): pass
    elif what.startswith('Withdrawal at '): pass
    elif what.startswith('Deposit'): what = what.replace('Deposit', '', 1).strip()
    elif what.startswith('Withdrawal'): what = what.replace('Withdrawal', '', 1).strip()
    what = what.replace('TYPE:', ' TYPE:').replace('  TYPE:', ' TYPE:')
    return what


def _categorizeFidelity(what):
    cat = sub_cat = ''
    if what.startswith('CONTRIBUTION'): cat, sub_cat = 'Transfer', 'Account'
    elif what.startswith('DIVIDEND'): cat, sub_cat = 'Income', 'Dividends'
    elif what.startswith('INTEREST'): cat, sub_cat = 'Income', 'Interest'
    return [cat, sub_cat]


def _categorizeVanguard(what):
    cat = sub_cat = ''
    if what.startswith('BUY ELEC'): cat, sub_cat = 'Transfer', 'Account'
    elif what.startswith('INCOME D'): cat, sub_cat = 'Income', 'Dividends'
    elif what.startswith('LT CAP G') or what.startswith('ST CAP G'): cat, sub_cat = 'Income', 'Gains'
    return [cat, sub_cat]


def _categorizeUMB(what):
    cat = sub_cat = ''
    if what.startswith('CURRENT YEAR CONTRIBUTION'): cat, sub_cat = 'Transfer', 'Account'
    return [cat, sub_cat]