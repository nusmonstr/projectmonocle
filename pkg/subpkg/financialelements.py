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
from datetime import datetime, timedelta, date
from math import floor
from openpyxl import load_workbook
import csv
from collections import namedtuple
from calendar import monthrange


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


def valid_dates(first_payperiod, valid_pay_periods):
    """ This function acts as a placeholder for a proper calculation of payperiods on semimonthly basis.
        """
    periods = []
    end_date = datetime.date(datetime.now())
    for period in valid_pay_periods:
        periods.append(date(*[int(x) for x in period.split('-')]))
    periods = set([x for x in periods if x > first_payperiod and x <= end_date])
    return periods


def midlast_month_weekdays(start_date_str):
    """ This function acts as a range() like operation for datetime objects.
        The function returns a *set* of datetime objects from provided start date to the present including
        only the 15th and last day of the month. This function returns weekdays only, and will search for
        the previous or next weekday depending on early_weekday boolean."""
    start_date = datetime.strptime(start_date_str, '%m-%d-%Y').date()
    end_date = datetime.date(datetime.now())
    start_pair = (start_date.year, start_date.month)
    stop_pair = (end_date.year, end_date.month)
    periods = set()
    # Figure out month/year combos between start and end
    comprehensive_pairs = [(year, month) for year in range(start_date.year, end_date.year+1) for month in range(1, 13) if (year, month)>=start_pair and (year, month) <= stop_pair]
    # Iterate through each and calculate 15th and last datetime obj
    for iyear, imonth in comprehensive_pairs:
        mid_nearest_workday = nearest_previous_workday(datetime(year=iyear, month=imonth, day=15)).date()
        if mid_nearest_workday >= start_date and mid_nearest_workday <= end_date:
            periods.add(mid_nearest_workday)
        last_nearest_workday = nearest_previous_workday(last_day_of_month(mid_nearest_workday))
        if last_nearest_workday >= start_date and last_nearest_workday <= end_date:
            periods.add(last_nearest_workday)
    return periods


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
    '''
    # Cleanup Citi Downloads because of dumb naming
    raw = input('Are you sure you downloaded Personal Citi Bank Data first?')
    if raw:
        exit('Go Fix Citi and rerun')
    citi_downloads = [x for x in directory_listing if x.startswith('Since') and x.endswith('.CSV')]
    citi_downloads = sorted([(x, datetime.fromtimestamp(os.path.getctime(x))) for x in citi_downloads], key=lambda x: x[1])
    [os.remove(x) for x in directory_listing if x.endswith('_CURRENT_VIEW.CSV')]
    os.rename(citi_downloads[0][0], 'MC_188_CURRENT_VIEW.CSV')
    os.rename(citi_downloads[1][0], 'MC_317_CURRENT_VIEW.CSV')
    '''
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
            if elements[0]: # Skip rows without a date
                pubdate, desc, amnt, cat, subcat, tag, note, status, taxation, paypstyle, payp, acnt, fund, added, *extra = elements
                record = Transaction(pubdate, desc, amnt, cat, subcat, tag, note, status, taxation, paypstyle, payp, acnt, fund, added)
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
                pubdate, desc, amnt, cat, subcat, tag, note, status, taxation, paypstyle, payp, acnt, fund, added, *extra = line
                pubdate = pubdate.replace('-', '') + 'Ymd'
                record = Transaction(pubdate, desc, amnt, cat, subcat, tag, note, status, taxation, paypstyle, payp, acnt, fund, added)
                statement.append(record)
    return statement


class Transaction:
    """ This class defines the basic fundamental record used for financial record keeping. """
    def __init__(self, pubdate='1/1/2016', desc='Default', amnt='0', cat='', subcat='', notes='', tag='', status='Default', taxation='', paypstyle = '', payp='', acnt='Default', fund='', added=''):
        self.desc = desc
        self.amnt = amnt
        self.cat = cat
        self.subcat = subcat
        self.notes = notes
        self.tag = tag
        self.status = status
        self.taxation = taxation
        self.paypstyle = paypstyle
        self.payp = ''
        self.acnt = acnt
        self.fund = fund
        self.added = added
        self.datenum = ''
        self.pubdate = pubdate  # This is last because it sets other attributes

    def get_pubdate(self):
        return self._pubdate

    def set_pubdate(self, value):
        #print(self.desc, value, type(value)) #For Troubleshooting Date setting
        if isinstance(value, str):
            if value.endswith('Ymd'):
                value = value[4:6]+'/'+value[6:8]+'/'+value[0:4]
            try:
                value = datetime.date(datetime.strptime(value, '%m/%d/%Y'))
            except:
                print('Failed to convert to date')
                print('description <{}>'.format(self.desc))
                print('amount <{}>'.format(self.amnt))
                print('date value <{}>'.format(value))
                exit()

        elif isinstance(value, datetime):
            value = datetime.date(value)
        self._pubdate = value
        self.payp = calculate_payperiod(value, self.paypstyle)
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
        return ', '.join([pubdatestr, (self.desc+' '*20)[:20], (str(self.amnt)+' '*8)[:8], self.cat, self.subcat, self.notes, self.tag, self.status, self.taxation, self.paypstyle, paypstr, self.acnt, self.fund, self.added, str(self.datenum)])

    def spill(self, empty_columns):
        return [self.pubdate, self.desc, self.amnt, self.cat, self.subcat, self.notes, self.tag, self.status, self.taxation, self.paypstyle, self.payp, self.acnt, self.fund, self.added, self.datenum] + [''] * empty_columns

    def spoon_feed(self):
        raw = [self.pubdate, self.desc, self.amnt, self.cat, self.subcat, self.notes, self.tag, self.status, self.taxation, self.paypstyle, self.payp, self.acnt, self.fund, self.added, self.datenum]
        return [str(x) for x in raw]


def calculate_payperiod(pubdate, frequency):
    """ This function returns the payperiod datetime object that a provided date belongs to."""
     #Payperiod In Use 	 Payperiod 	 Semimonthly 	 Biweekly Up 	 Biweekly Down
    if frequency == 'Semimonthly':
        # for may 18th 2009, what is the previously past billing
        # step one, figure out if >= last weekday of month
        last_weekday_of_month = nearest_previous_workday(last_day_of_month(pubdate))
        if pubdate >= last_weekday_of_month:
            return last_weekday_of_month
        else:
            # if yes then yes else calculate the half date
            month_midpoint = nearest_previous_workday(pubdate.replace(day=15))
            # step two, which side? if greater then half else last day of last month
            if pubdate >= month_midpoint:
                return month_midpoint
            else:
                last_day_of_previous_month = pubdate.replace(day=1) - timedelta(days=1)
                last_weekday_of_previous_month = nearest_previous_workday(last_day_of_previous_month)
                return last_weekday_of_previous_month
    elif frequency.startswith('Biweekly'):
        first_payperiod = datetime.date(datetime(1999, 1, 1))   # hard coded dates just used for offset, arbitrarily called Up and Down
        if frequency.endswith('Down'):
            first_payperiod = datetime.date(datetime(1999, 1, 8))
        days_gone = pubdate - first_payperiod
        periods = floor(days_gone.days/14)
        offset = timedelta(days=periods*14)
        return first_payperiod + offset
    # Options include Monthly first weekday, last weekday, arbitrary nearest  prior weekday, arbitrary nearest next weekday
    elif frequency.startswith('Monthly'):
        pass

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


def is_weekday(any_date):
    if any_date.weekday() <= 4:
        return True
    else:
        return False


def nearest_previous_workday(any_date):
    nearest = False
    while not nearest:
        if is_weekday(any_date):
            nearest = any_date
        else:
            any_date += timedelta(days = -1)
    return nearest


def last_day_of_month(anydate):
    return anydate.replace(day=monthrange(anydate.year, anydate.month)[1])

if __name__ == "__main__":

    for each in midlast_month_weekdays(date(year=2010, month=8, day=6)):
        print(each)
    exit(0)
    with open('pubdates.txt', 'r') as fh:
        raw_dates = fh.readlines()
    all_datetimes = [datetime.strptime(x.strip(), '%m/%d/%Y') for x in raw_dates]
    '''
    with open('semiperiod.txt', 'w') as ans:
        all_datetimes = [datetime.strptime(x.strip(), '%m/%d/%Y') for x in raw_dates]
        for each in all_datetimes:
            ans.write(calculate_payperiod(each, 'semi').strftime('%m/%d/%Y')+'\n')
    exit(0)
    '''
    for each in all_datetimes[4000:4020]:
        print('For date of', each)
        print('\tSemi period is', calculate_payperiod(each,'Semimonthly'))
        print('\tBiweekly Up period is', calculate_payperiod(each.date(),'Biweekly Up'))
        print('\tBiweekly Down period is', calculate_payperiod(each.date(),'Biweekly Down'))


'''
Friday, January 13, 2017
Tuesday, January 31, 2017
Wednesday, February 15, 2017
Tuesday, February 28, 2017
Wednesday, March 15, 2017
Friday, March 31, 2017
Friday, April 14, 2017
Friday, April 28, 2017
Monday, May 15, 2017
Wednesday, May 31, 2017
Thursday, June 15, 2017
Friday, June 30, 2017
Friday, July 14, 2017
Monday, July 31, 2017
Tuesday, August 15, 2017
Thursday, August 31, 2017
Friday, September 15, 2017
Friday, September 29, 2017
Friday, October 13, 2017
Tuesday, October 31, 2017
Wednesday, November 15, 2017
Thursday, November 30, 2017
Friday, December 15, 2017
Friday, December 29, 2017
'''
