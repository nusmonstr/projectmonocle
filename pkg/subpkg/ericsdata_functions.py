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

userpayp = 'Semimonthly'
autocat_filepath = r'C:\Users\erchampe\OneDrive\Python\Github\ProjectMonocle\local\autocat_strings.csv'

autocat_string_dict = dict()
with open(autocat_filepath, 'r') as fh:
    raw = [x for x in fh.readlines() if x]
    for record in raw:
        record = [x.replace('<comma>', ',').strip() for x in record.split(',')]
        key, *value = record
        autocat_string_dict[key] = tuple(value)
autocat_strings = sorted(autocat_string_dict.keys(), key=lambda x: len(x), reverse=True)


vanguard_accounts = {
    '73510370': 'Vanguard',
    '46069832': 'Vanguard - Roth Brokerage',
    '53949785': 'Vanguard - IRA Brokerage'
}

vanguard_descriptions = {
    'Buy': ('Transfer', ''),
    'Buy using the Proceeds from the Sale of Another Fund': ('Transfer', 'Fund'),
    'Dividend Received': ('Income', 'Dividends'),
    'Dividend Reinvestment': ('Income', 'Dividends'),
    'Funds received via Electronic Bank Transfer': ('Transfer', 'Account'),
    'Incoming Mutual Fund Share Class Conversion': ('Transfer', 'Fund'),
    'Long-Term Capital Gains Distribution': ('Income', 'Gains'),
    'Outgoing Mutual Fund Share Class Conversion': ('Transfer', 'Fund'),
    'Reinvestment of a Long-Term Capital Gains Distribution': ('Income', 'Gains'),
    'Reinvestment of a Short-Term Capital Gains Distribution': ('Income', 'Gains'),
    'Rollover Contribution': ('Transfer', 'Account'),
    'Rollover Conversion': ('Transfer', 'Account'),
    'Sell in Order to Buy Shares of Another Fund': ('Transfer', 'Fund'),
    'Short-Term Capital Gains Distribution': ('Income', 'Gains'),
    'Sweep Into Money Market Settlement Fund': ('Transfer', ''),
    'Sweep Out Of Money Market Settlement Fund': ('Transfer', 'Fund'),
    'Total Conversion under 59 1/2': ('Transfer', 'Account')
}

fidelity_descriptions = {
    'Change in Market Value': ('Income', 'Valuation'),
    'Change on Market Value': ('Income', 'Valuation'),
    'CONTRIBUTION': ('Transfer', 'Account'),
    'DIVIDEND': ('Income', 'Dividends'),
    'Dividends': ('Income', 'Dividends'),
    'Exchange In': ('Transfer', 'Fund'),
    'Exchange Out': ('Transfer', 'Fund'),
    'INTEREST': ('Income', 'Interest')
}

umb_descriptions = {
    'BROKERAGE DEBIT': ('Transfer', 'Account'),
    'Health Savings Purchase Manual': ('Transfer', 'Account'),
    'INTEREST': ('Income', 'Interest'),
    'WEB-INITIATED ACH DISBURSEMENT': ('Transfer', 'Account'),
    'CURRENT YEAR CONTRIBUTION': ('Transfer', 'Account'),
    'EMPLOYER CURRENT YEAR CONTRIBUTI': ('Transfer', 'Account'),
    'TRUSTEE TO TRUSTEE TRANSFER': ('Transfer', 'Account')
}

vanguard_amount_multipliers = {
    'Buy': -1,
    'Buy using the Proceeds from the Sale of Another Fund': -1,
    'Dividend Received': 1,
    'Dividend Reinvestment': 0,
    'Funds received via Electronic Bank Transfer': 0,
    'Incoming Mutual Fund Share Class Conversion': 1,
    'Long-Term Capital Gains Distribution': 1,
    'Outgoing Mutual Fund Share Class Conversion': 1,
    'Reinvestment of a Long-Term Capital Gains Distribution': 0,
    'Reinvestment of a Short-Term Capital Gains Distribution': 0,
    'Rollover Contribution': 1,
    'Rollover Conversion': 1,
    'Sell in Order to Buy Shares of Another Fund': -1,
    'Short-Term Capital Gains Distribution': 1,
    'Sweep Into Money Market Settlement Fund': -1,
    'Sweep Out Of Money Market Settlement Fund': -1,
    'Total Conversion under 59 1/2': 1
}

vanguard_fund_lookup_1 = {
    'CASH': 'Federal Money Market',
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
    'VANGUARD LONG TERM TREASURY INDEX ADMIRAL CL': 'Longterm Gov Bond Index',
    'VANGUARD SHORT TERM BOND INDEX ADMIRAL CL': 'Short-Term Bond Index Adm',
    'VANGUARD GLOBAL CAPITAL CYCLES INVESTOR CL': 'Capital Cycles?'}

vanguard_fund_lookup_2 = {
    'Vanguard 500 Index Fund Admiral Shares': '500 Index Fund Adm',
    'Vanguard Short-Term Bond Index Fund Admiral Shares': 'Short-Term Bond Index Adm',
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

fidelity_fund_lookup = {
    'SPAXX': 'Fidelity Gov Money Mkt',
    'MSFT': 'Microsoft Corp'
}


def import_brokerage(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, skip_header, None):
            if line:
                stripped_line = [x.strip() for x in line]
                pubdate, desc, symbol, fund, type, shares, price, commission, fee, acc_int, amnt, *trash = stripped_line
                if amnt.startswith('-'):
                    amnt = amnt[1:]
                elif amnt:
                    amnt = '-' + amnt
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
                s.cat, s.subcat = _autocategorize(desc)
                s.notes = 'Share Price: {} Shares: {} Commission: {} Fees: {}'.format(price, shares, commission, fee)
                s.fund = symbol
                s.tag = ''
                statement.append(s)
            else:
                return statement, list()


def import_citi(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, skip_header, None):
            if line:
                state, pubdate, desc, deb, cred, *trash = line
                cred = cred.replace('-', '')
                amnt = '-' + deb if deb is not '' else cred # After Citi change to - on deb used to be == '-' + deb if deb is not '' else cred
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
                s.cat, s.subcat = _autocategorize(desc)
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
                s.cat, s.subcat = fidelity_descriptions[desc]
                s.notes = 'Shares/Unit '+shares+' | '+fund
                s.tag = ''
                statement.append(s)
    return statement, list()


def import_firsttech(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, skip_header, None):
            if line:
                id, pubdate, effdate, type, amnt, check_no, ref, desc, tran_cat, method, balance = line
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
                s.cat, s.subcat = _autocategorize(desc)
                s.notes = ''
                s.tag = ''
                statement.append(s)
                #values.append((pubdate, balance))
    # Balance Gap is not working, amount is not accurate $$$
    #baldate = datetime.date(datetime.fromtimestamp(os.path.getmtime(bank_data)))
    #balance = round(float(sorted(values, key=lambda x: x[0], reverse=True)[0][1]), 2)
    #v = Transaction(baldate, 'ACCOUNT BALANCE GAP', balance, paypstyle=userpayp, acnt=bank_name, status=bank_status)
    return statement, values


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
                s.cat, s.subcat = umb_descriptions[desc]
                s.notes = ''
                s.tag = ''
                statement.append(s)
    return statement, list()


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
                s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=vanguard_accounts[account], status=bank_status)
                s.cat, s.subcat = vanguard_descriptions[desc]
                s.notes = 'Share Price: {} Shares: {}'.format(price, shares)
                s.fund = vanguard_fund_lookup_1[fund]
                if vanguard_amount_multipliers[desc] == 0:
                    s.notes = s.notes + ' Value: ' + str(s.amnt)
                    s.amnt = 0
                s.amnt = s.amnt * vanguard_amount_multipliers[desc]
                statement.append(s)
    pubdate = datetime.date(datetime.fromtimestamp(os.path.getmtime(bank_data)))
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, val_start, trans_start-1):
            # Verify that the line is not empty before unpacking
            if ''.join(line):
                account, desc, symbol, shares, price, amnt, *trash = line
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=vanguard_accounts[account], status=bank_status)
                s.cat, s.subcat = ('Income', 'Valuation')
                s.notes = 'Share Price: {} Shares: {}'.format(price, shares)
                s.fund = vanguard_fund_lookup_2[desc]
                values.append(s)
    return statement, values


def import_widget(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    keys = ['DTPOSTED', 'MEMO', 'TRNAMT', 'NAME']
    fh = open(bank_data, 'r')
    lines = fh.readlines()
    lines = ''.join(lines).replace('\n', '')
    # for each in lines: print(each)
    #line_Val = findall(r'<LEDGERBAL><BALAMT>(.*?)</LEDGERBAL>', lines)
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
        s.cat, s.subcat = _autocategorize(desc)
        s.notes = ''
        s.tag = ''
        statement.append(s)
    #baldate = datetime.date(datetime.fromtimestamp(os.path.getmtime(bank_data)))
    #balance = line_Val[0].replace('<', ',').split(',')[0]
    #v = Transaction(baldate, 'ACCOUNT BALANCE GAP', balance, paypstyle=userpayp, acnt=bank_name, status=bank_status)
    #values.append(v)
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

def _autocategorize(description):
    for strstart in autocat_strings:
        if description.startswith(strstart):
            return autocat_string_dict[strstart]
    return ('', '')