"""
Created on Aug 23, 2016
Modified on Dec 11, 2016

@author: Eric Champe

Functions:
Transaction
    Bank namedtuple
    find_missing
    periodic_dates
    archive_cleanup
    find_downloads_today
    unpack_transactions
    Transaction class
    calculate_payperiod
Helper Functions
    write_to_csv
    batch_setattr
    datetime_to_xl
    nones_to_null

"""
import os
from datetime import datetime, timedelta
from math import floor
from openpyxl import load_workbook
import csv
from collections import namedtuple


Bank = namedtuple('Bank', 'filename title extraction header status')
info = '    [!] '
question = '[?] '
response = '\t[select]: '


def find_missing(candidate_list, base_list):
    """ This function compares candidates against an existing comprehensive database
        and returns the records that do not have a three field match with the base"""
    missing = list()
    for candidate in candidate_list:
        for record in base_list:
            if candidate.pubdate == record.pubdate:
                if candidate.desc == record.desc:
                    if candidate.amnt == record.amnt:
                        break   # A match was found, break to look for match with next candidate
        else:
            missing.append(candidate)
    return missing


def periodic_dates(first_payperiod, payperiod_length):
    """ This function acts as a range() like operation for datetime objects.
        The function returns a *set* of datetime objects from provided start date to the present with a provided
        step duration in days."""
    end_date = datetime.date(datetime.now())
    days_past = int((end_date - first_payperiod).days)
    periods = set(first_payperiod + timedelta(days=period) for period in range(0, days_past + 1, payperiod_length))
    return periods


def archive_cleanup(archive_filename, copies_to_keep=5, archive_path=None):
    """ This function cleans out old backups."""
    original_path = os.getcwd()
    if archive_path:
        os.chdir(archive_path)
    directory_listing = [file for file in os.listdir(os.getcwd())]
    name, extension = archive_filename.split('.')
    # Find all files that have the expected filename and extension regardless of name suffix
    related_downloads = [file for file in directory_listing if (file.startswith(name+'_') and file.endswith(extension))]
    related_downloads.sort(reverse=True)
    for file in related_downloads[copies_to_keep:]:
        os.remove(file)
    os.chdir(original_path)


def find_downloads(full_banklist, download_path=None):
    """ This function returns a list of newly downloaded filenames and deletes older files.
    The function will also print the total number of downloads that will be included in this update."""
    original_path = os.getcwd()
    if download_path:
        os.chdir(download_path)
    directory_listing = [file for file in os.listdir(os.getcwd())]
    available_banks = list()
    for bank in full_banklist:
        name, extension = bank.filename.split('.')
        # Find all files that have the expected filename and extension regardless of name suffix
        related_downloads = [file for file in directory_listing if (file.startswith(name) and file.endswith(extension))]
        if related_downloads:
            related_downloads = sorted(related_downloads, key=lambda x: datetime.fromtimestamp(os.path.getctime(x)), reverse=True)
            available_banks.append(Bank(related_downloads[0], bank[1], bank[2], bank[3], bank[4]))
            # Delete all but the newest bank download for this bank
            for file in related_downloads[1:]:
                #print(question + '{} is not the most recent {} bank download, [Enter] to delete, [k] to keep.'.format(file,bank[1]))
                #user_response = input(response)
                user_response = 'delete'
                if user_response.lower().strip() != 'k':
                    os.remove(file)
                else:
                    print('Keeping {}'.format(file))
    total_downloads = len(available_banks)
    found_msg = info + '{} downloads will be included in this update.'.format(total_downloads)
    absent_msg = info + 'No downloads were found in {}\n[x] to exit,[any] to continue...\n>>> '.format(download_path)
    if total_downloads:
        print(found_msg)
    elif str(input(absent_msg)).lower() == 'x':
        exit()
    os.chdir(original_path)
    return available_banks


def unpack_transactions(filepath, sheet_name='', origin=None):
    """ This wrapper function calls appropriate function for file type provided."""
    if filepath.endswith('.csv'):
        return unpack_transactions_csv(filepath)
    elif filepath.endswith('.xlsx'):
        return unpack_transactions_xl(filepath, sheet_name, origin)
    else:
        raise ValueError('The filepath provided does not a valid .xlsx or .csv file.')


def unpack_transactions_xl(workbook_name, sheet_name, origin):
    statement = list()
    workbook = load_workbook(filename=workbook_name, read_only=True)
    worksheet = workbook[sheet_name]
    skip_rows = max(origin[0]-1, 0)
    skip_cols = max(origin[1]-1, 0)
    for row in worksheet.rows:
        if skip_rows > 0:
            skip_rows -= 1  # Skip past next row until the origin is reached
        else:
            elements = [cell.value for cell in row][skip_cols:] # Skip to the origin column
            if None in elements:
                elements = nones_to_null(elements)   # Swap None for empty strings
            pubdate, desc, amnt, cat, subcat, tag, note, status, payp, acnt, added, *extra = elements
            record = Transaction(pubdate, desc, amnt, cat, subcat, tag, note, status, payp, acnt, added)
            statement.append(record)
    return statement


def unpack_transactions_csv(csv_name):
    statement = list()
    with open(csv_name) as f:
        lines = csv.reader(f)
        next(lines)    # Skip header row
        for line in lines:
            if line:    # Verify that the line is not empty before unpacking
                if None in line:
                    line = nones_to_null(line)   # Swap None for empty strings
                pubdate, desc, amnt, cat, subcat, tag, note, status, payp, acnt, added, *extra = line
                pubdate = pubdate.replace('-', '') + 'Ymd'
                record = Transaction(pubdate, desc, amnt, cat, subcat, tag, note, status, payp, acnt, added)
                statement.append(record)
    return statement


class Transaction:
    """ This class defines the basic fundamental record used for financial record keeping. """
    def __init__(self, pubdate='1/1/2016', desc='Default', amnt='0', cat='', subcat='', notes='', tag='', status='Default', payp='', acnt='Default', added=''):
        self.desc = desc
        self.amnt = amnt
        self.cat = cat
        self.subcat = subcat
        self.notes = notes
        self.tag = tag
        self.status = status
        self.payp = ''
        self.acnt = acnt
        self.added = added
        self.datenum = ''
        self.pubdate = pubdate  # This is last because it sets other attributes

    def get_pubdate(self):
        return self._pubdate

    def set_pubdate(self, value):
        if isinstance(value, str):
            if value.endswith('Ymd'):
                value = value[4:6]+'/'+value[6:8]+'/'+value[0:4]
            value = datetime.date(datetime.strptime(value, '%m/%d/%Y'))
        elif isinstance(value, datetime):
            value = datetime.date(value)
        self._pubdate = value
        self.payp = calculate_payperiod(value)
        self.datenum = datetime_to_xl(value)
    pubdate = property(get_pubdate, set_pubdate)

    def get_amnt(self):
        return self._amnt

    def set_amnt(self, value):
        if isinstance(value, str):
            value = value.replace(',', '').replace('$', '').replace('(', '').replace(')', '').strip()
            # Evaluate excel formula strings to a numeric
            if value.startswith('='):
                value = eval(value[1:])
            elif value.replace(' ', '') == '':
                value = '999999999'         # This was added when one of bank activity files had transaction without an amount... $0 was blank
        self._amnt = round(float(value), 2)
    amnt = property(get_amnt, set_amnt)

    def get_desc(self):
        return self._desc

    def set_desc(self, value):
        self._desc = str(value).strip()
    desc = property(get_desc, set_desc)

    def get_notes(self):
        return self._notes

    def set_notes(self, value):
        self._notes = str(value).strip()
    notes = property(get_notes, set_notes)

    def get_tag(self):
        return self._tag

    def set_tag(self, value):
        self._tag = str(value).strip()
    tag = property(get_tag, set_tag)

    def __repr__(self):
        pubdatestr = datetime.strftime(self.pubdate,'%m/%d/%y')
        paypstr = datetime.strftime(self.payp, '%m/%d/%y')
        return ', '.join([pubdatestr, (self.desc+' '*20)[:20], (str(self.amnt)+' '*8)[:8], self.cat, self.subcat, self.notes, self.tag, self.status, paypstr, self.acnt, self.added, str(self.datenum)])

    def spill(self, empty_columns):
        return [self.pubdate, self.desc, self.amnt, self.cat, self.subcat, self.notes, self.tag, self.status, self.payp, self.acnt, self.added, self.datenum] + [''] * empty_columns

    def spoon_feed(self):
        raw = [self.pubdate, self.desc, self.amnt, self.cat, self.subcat, self.notes, self.tag, self.status, self.payp, self.acnt, self.added, self.datenum]
        return [str(x) for x in raw]

def calculate_payperiod(pubdate):
    """ This function returns the payperiod datetime object that a provided date belongs to."""
    first_payperiod = datetime.date(datetime(2016, 1, 8))   # Need to remove this hardcoded first paycheck $$$
    if pubdate < first_payperiod:
        first_payperiod = datetime.date(datetime(2005, 1, 1))
    days_gone = pubdate - first_payperiod
    periods = floor(days_gone.days/14)
    offset = timedelta(days=periods*14)
    return first_payperiod + offset


#########################################################################
#   Generic Functions


def write_to_csv(csv_name, records, field_names):
    """ This function uses the csv module to write a csv data file, complete
        with a header and all records"""
    with open(csv_name, 'w', newline='') as file_handle:
        writer = csv.writer(file_handle, field_names)
        writer.writerow(field_names)
        for record in records:
            writer.writerow(record.spill(0))


def batch_setattr(my_list, attr_name, value):
    """ This function takes a list of objects and assigns value to every object's attribute.
        my_list[:].attr_name = value """
    for obj in my_list:
        setattr(obj, attr_name, value)


def datetime_to_xl(datetime_object):
    """ This function converts a datetime object to an integer for MS Excel."""
    xl2py = 693594
    pyint = datetime_object.toordinal()
    xl_int = pyint-xl2py
    return xl_int


def nones_to_null(my_list):
    """ This function replaces instances of None in a list with null strings, ''."""
    no_holes = list()
    for x in my_list:
        if x is None:
            no_holes.append('')
        else:
            no_holes.append(x)
    return no_holes
