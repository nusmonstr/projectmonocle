"""
Created on Aug 23, 2016

@author: Eric Champe

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
import win32com.client


userpayp = 'Semimonthly'


fund_lookup_adp = {
    '74149P705': 'T. Rowe Price Retirement 2020 Fund - Class R',
    '09251M108': 'BlackRock Equity Dividend Fund - Investor A Class',
    '85744A513': 'State Street S+P 500 Index Securities Lending Series Fund - Class IX',
    '92647Q587': 'Victory RS Large Cap Alpha Fund - Class  A',
    '85744A471': 'State Street S+P MidCap Index Non-Lending Series Fund - Class J',
    '85744A554': 'State Street Russell Small Cap Index Securities Lending Series Fund - Class VIII',
    '85744A596': 'State Street International Index Securities Lending Series Fund - Class VIII'}


fund_lookup_hsainv = {
    'VBMPX': 'Vanguard Total Bond Market Institutional Plus Index',
    'VGSNX': 'Vanguard REIT Index I',
    'VIIIX': 'Vanguard Institutional Plus Index',
    'VWIAX': 'Vanguard Wellesley Income Admiral'}


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
    values = list()
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, skip_header, None):
            # Verify that the line is not empty before unpacking
            if line:
                state, pubdate, desc, deb, cred, *trash = line
                amnt = '-' + deb if deb is not '' else cred
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
                s.cat = s.subcat = ''
                s.notes = ''
                s.tag = ''
                statement.append(s)
    return statement, values


def import_umbbank(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, skip_header, None):
            # Verify that the line is not empty before unpacking
            if line:
                pubdate, trans_type, desc, memo, amnt = line
                desc = desc + memo.strip()
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
                s.cat = s.subcat = ''
                s.notes = ''
                s.tag = ''
                statement.append(s)
    return statement, values


def import_fidelity(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, skip_header, None):
            # Verify that the line is not empty before unpacking
            if line:
                pubdate, fund, desc, amnt, shares, *trash = line
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
                s.cat, s.subcat = _categorizeFidelity(s.desc)
                s.notes = 'Shares/Unit '+shares+' | '+fund
                s.tag = ''
                statement.append(s)
    return statement, values


def import_firsttech(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, skip_header, None):
            # Verify that the line is not empty before unpacking
            if line:
                id, pubdate, effdate, type, amnt, check_no, ref, desc, *trash = line
                # Create new transaction object with formatted data
                s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
                s.cat = s.subcat = ''
                s.notes = ''
                s.tag = ''
                statement.append(s)
    return statement, values


def import_healthequity(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    fh = open(bank_data, 'r')
    lines = fh.readlines()
    # $$$ Make this into a function for easier troubleshooting - Convert qif formatting to csv formatting
    lines = ''.join(lines).replace('\nT', '\t').replace('\nM', ',').replace('\nD', ',').replace('\n^', '').replace('\n',
                                                                                                                   '').split(
        '\t')
    for line in lines[skip_header:]:
        # Verify that the line is not empty before unpacking
        if [line]:
            amnt, pubdate, desc, *trash = line.split(',')

            # Create new Transaction object with formatted data
            s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
            s.cat, s.subcat = _categorizeHealthEquity(s.desc)
            s.notes = ''
            s.tag = ''
            statement.append(s)
    return statement, values


def import_healthequityinv(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    download_path = os.getcwd()
    bank_data_filepath = os.path.join(download_path, bank_data)
    wb = EasyExcel(bank_data_filepath)
    sheet_name = bank_data.split('.', 1)[0][:31]
    cell_range = wb.getContiguousRange(sheet_name, 4, 1)
    for line in cell_range:
        # Verify that the line is not empty before unpacking
        if [line]:
            pubdate, fund, cat, memo, price, amnt, share, shares, value = line

            # Create new Transaction object with formatted data
            desc = ' '.join([memo, cat])
            price, share = [str(x) for x in [price, share]]
            s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
            s.cat, s.subcat = _categorizeHealthEquityInv(cat)
            s.notes = 'Share Price: ' + price + ' Shares: ' + share
            s.fund = fund_lookup_hsainv[fund]
            statement.append(s)
    wb.close()
    return statement, values


def import_widget(bank):
    # Need to pull both checking Transactions then repeat for savings
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    keys = ['DTPOSTED', 'MEMO', 'TRNAMT', 'NAME']
    fh = open(bank_data, 'r')
    lines = fh.readlines()
    lines = ''.join(lines).replace('\n', '')
    # for each in lines: print(each)
    line_Val = findall(r'<LEDGERBAL>(.*?)</LEDGERBAL>', lines)
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
        s.cat = s.subcat = ''
        s.notes = ''
        s.tag = ''
        statement.append(s)
    baldate = datetime.date(datetime.fromtimestamp(os.path.getmtime(bank_data)))
    balance = line_Val[0].replace('<', ',').replace('>', ',').split(',')
    v = Transaction(baldate, 'Placeholder', balance[2], paypstyle=userpayp, acnt=bank_name, status=bank_status)
    values.append(v)
    return statement, values


def import_adp(bank):
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    fh = open(bank_data, 'r')
    lines = fh.readlines()
    lines = ''.join(lines).replace('\n', '')
    lines_div = findall(r'<REINVEST>(.*?)</REINVEST>', lines)  # Grab Dividend and Earning lines
    lines_val = findall(r'<INVPOS>(.*?)</INVPOS>', lines)  # Grab Market Value lines
    lines = findall(r'<INVTRAN>(.*?)</BUYMF>', lines)  # Trim all but Contribution lines
    lines.extend(lines_div)  # Add Dividend and Earning lines to Contribution lines
    keys = ['DTSETTLE', 'MEMO', 'UNITPRICE', 'UNIQUEID', 'UNITS', 'TOTAL']  # pubdate, desc, price, fundcode, shares and amnt respectively
    for line in lines:
        line = line[1:].replace('<', ',').replace('>', ',').split(',')
        tags, targets = line[::2], line[1::2]  # Return tags and values in an alternating sequence
        pairs = dict(zip(tags, targets))
        pubdate, desc, price, fundcode, shares, amnt = [pairs.get(key) for key in keys]
        pubdate += 'Ymd'
        amnt = amnt[1:] if amnt.startswith('-') else '-' + amnt # reverse the sign on the amount string
        # Create new Transaction object with formatted data
        s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
        s.cat, s.subcat = _categorizeADP(desc)
        s.notes = 'Share Price: ' + price + ' Shares: ' + shares
        s.fund = fund_lookup_adp[fundcode]
        statement.append(s)
    pubdate = datetime.date(datetime.fromtimestamp(os.path.getmtime(bank_data)))
    desc = 'Syncronize my 401K Value'
    keys = ['UNITS', 'UNITPRICE', 'MKTVAL', 'DTPRICEASOF', 'UNIQUEID']  # shares, price, amnt, when and fundcode respectively
    for line in lines_val:
        line = line[6:].replace('<', ',').replace('>', ',').split(',')
        tags, targets = line[::2], line[1::2]  # Return tags and values from an alternating sequence
        pairs = dict(zip(tags, targets))
        shares, price, amnt, when, fundcode = [pairs.get(key) for key in keys]
        # Create new Transaction object with formatted data
        s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
        s.cat, s.subcat = 'Income', 'Valuation'
        s.notes = 'Share Price: ' + price + ' Shares: ' + shares
        s.fund = fund_lookup_adp[fundcode]
        values.append(s)
    return statement, values


def import_vanguard(bank):
    # Vanguard Bank #############################
    skip_header, bank_data, bank_name, bank_status = bank.header, bank.filename, bank.title, bank.status
    statement = list()
    values = list()
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, skip_header, None):
            # Verify that the line is not empty before unpacking
            if line: 
                account, when0, pubdate, trantype, desc, fund, symbol, shares, price, amnt, *trash = line
                # Create new transaction object with formatted data
                trash = '='.join(trash)
                line_is = [account, when0, pubdate, trantype, desc, fund, symbol, shares, price, amnt, trash]
                #print('|'.join(line_is))
                s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
                s.cat, s.subcat = _categorizeVanguard(desc)
                s.notes = 'Share Price: '+price+' Shares: '+shares
                s.fund = vanguard_fund_lookup_1[fund]
                statement.append(s)
    pubdate = datetime.date(datetime.fromtimestamp(os.path.getmtime(bank_data)))
    with open(bank_data) as f:
        lines = reader(f)
        for line in islice(lines, 1, None):
            # Verify that the line is not empty before unpacking
            if line:
                if line[0].startswith('Account Number'):
                    break
                else:
                    account, desc, symbol, shares, price, amnt, *trash = line
                    # Create new transaction object with formatted data
                    #print('act:', account, type(account))
                    if account == '73510370':
                        s = Transaction(pubdate, desc, amnt, paypstyle=userpayp, acnt=bank_name, status=bank_status)
                        s.cat, s.subcat = 'Income','Valuation'
                        s.notes = 'Share Price: '+price+' Shares: '+shares
                        s.fund = vanguard_fund_lookup_2[desc]
                        values.append(s)
    return statement, values


def _widgetwhat(what,name):
    if what == '' : what = name if name != '' else 'Description not Found'
    elif what == 'Deposit' or what == 'Withdrawal' : pass
    elif what.startswith('Deposit at '): pass
    elif what.startswith('Deposit by '): pass
    elif what.startswith('Withdrawal at '): pass
    elif what.startswith('Deposit'): what = what.replace('Deposit','',1).strip()
    elif what.startswith('Withdrawal'): what = what.replace('Withdrawal','',1).strip()
    what = what.replace('TYPE:',' TYPE:').replace('  TYPE:',' TYPE:')
    return what


def _categorizeADP(what):
    cat = subcat = ''
    if what.startswith('Contribu') : cat,subcat = 'Transfer','Account'
    elif what.startswith('Dividend') : cat,subcat = 'Income','Dividends'
    return [cat,subcat]


def _categorizeFidelity(what):
    cat = subcat = ''
    if what.startswith('CONTRIBUTION') : cat,subcat = 'Transfer','Account'
    elif what.startswith('DIVIDEND') : cat,subcat = 'Income','Dividends'
    elif what.startswith('INTEREST') : cat,subcat = 'Income','Interest'
    return [cat,subcat]


def _categorizeHealthEquity(what):
    cat = subcat = ''
    if what.startswith('Employer') : cat,subcat = 'Income','Employer Contribution'
    elif what.startswith('Employee') : cat,subcat = 'Transfer','Account'
    elif what.startswith('Interest') : cat,subcat = 'Income','Interest'
    return [cat,subcat]


def _categorizeHealthEquityInv(what):
    cat = subcat = ''
    if what.startswith('Buy') : cat,subcat = 'Transfer', 'Account'
    elif what.startswith('Dividend') : cat,subcat = 'Income', 'Dividends'
    return [cat, subcat]


def _categorizeVanguard(what):
    cat = subcat = ''
    if what.startswith('BUY ELEC') : cat,subcat = 'Transfer','Account'
    elif what.startswith('INCOME D') : cat,subcat = 'Income','Dividends'
    elif what.startswith('LT CAP G') or what.startswith('ST CAP G') : cat,subcat = 'Income','Gains'
    return [cat,subcat]  


#########################################################################
#   File Functions


class EasyExcel:
    """A utility to make it easier to get at Excel.  Remembering
    to save the data is your problem, as is  error handling.
    Operates on one workbook at a time."""

    def __init__(self, filename=None):
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        if filename:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

    def save(self, newfilename=None):
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()

    def close(self):
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp

    def getCell(self, sheet, row, col):
        "Get value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Cells(row, col).Value

    def setCell(self, sheet, row, col, value):
        "set value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Value = value

    def getRange(self, sheet, row1, col1, row2, col2):
        "return a 2d array (i.e. tuple of tuples)"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value

    def setRange(self, sheet, leftCol, topRow, data):
        """insert a 2d array starting at given location.
        Works out the size needed for itself"""

        bottomRow = topRow + len(data) - 1
        rightCol = leftCol + len(data[0]) - 1
        sht = self.xlBook.Worksheets(sheet)
        sht.Range(
            sht.Cells(topRow, leftCol),
            sht.Cells(bottomRow, rightCol)
        ).Value = data

    def getContiguousRange(self, sheet, row, col):
        """Tracks down and across from top left cell until it
        encounters blank cells; returns the non-blank range.
        Looks at first row and column; blanks at bottom or right
        are OK and return None witin the array"""

        sht = self.xlBook.Worksheets(sheet)

        # find the bottom row
        bottom = row
        while sht.Cells(bottom + 1, col).Value not in [None, '']:
            bottom = bottom + 1
        # right column
        right = col
        while sht.Cells(row, right + 1).Value not in [None, '']:
            right = right + 1

        return sht.Range(sht.Cells(row, col), sht.Cells(bottom, right)).Value

    def fixStringsAndDates(self, aMatrix):
        # converts all unicode strings and times
        newmatrix = []
        for row in aMatrix:
            newrow = []
            for cell in row:
                if type(cell) is UnicodeType:
                    newrow.append(str(cell))
                elif type(cell) is TimeType:
                    newrow.append(int(cell))
                else:
                    newrow.append(cell)
            newmatrix.append(tuple(newrow))
        return newmatrix