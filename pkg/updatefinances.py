'''
Created on Jul 29, 2016
Modified on Dec 10, 2016
@author: Eric Champe
@coauthor: Mark Champe
'''


from .subpkg.ericsdata_functions import *
from .subpkg.marksdata_functions import *
from .subpkg.financialelements import unpack_transactions, batch_setattr, Bank, write_to_csv, find_missing, periodic_dates, find_downloads, archive_cleanup
from datetime import date, datetime
from copy import deepcopy
import win32com.client
import shutil
from sys import stdout
import os

def never_used():   # How do I tell pycharm that the functions are being used, I call the bank functions programmatically with extraction_fn
    import_capitalone()
    import_adp()


def main(args):

    # -- Fetch Online Bank Data -----------------------------------------
    # print('>>> Fetching bank data from the web...')
    # ***Not yet implemented
    # -- End Fetch Online Bank Data -------------------------------------

    print('args', args)

    # -- Setup ----------------------------------------------------------
    print('>>> Gathering downloads, payroll definition, and existing records...')
    # Variable declaration
    info = '    [!] '
    question = '[?] '
    response = '\t[select]: '

    users = dict()
    #print(os.getcwd(), os.listdir(os.getcwd()))
    config_path = os.path.dirname(os.path.abspath(__file__))
    existing_dir = os.getcwd()
    os.chdir(config_path)
    user_profiles = [file for file in os.listdir() if file.endswith('ini')]
    user_names = [file.replace('config_', '').replace('.ini', '') for file in user_profiles]
    for i, profile in enumerate(user_names):
        #print('\t{} {}'.format(i, profile))
        users[profile] = i
    #users = {'eric':0, 'mark':1, 'alex':3}
    if len(args) > 1 and args[1] in users.keys():
        selection = users[args[1]]  # TODO This is stupid; need to cleanup, don't need dict
    else:
        # Import User Configuration
        print(question + 'Which user profile would you like to use?')
        for i, profile in enumerate(user_profiles):
            print('\t{} {}'.format(i, profile))
        selection = int(input(response))
    print('You chose {}'.format(selection))

    bank_info = []
    with open(user_profiles[selection], 'r') as config:
        lines = config.readlines()
        user_dict = dict()
        for line in lines:
            line = line.strip()
            if line.startswith('<Bank>'):
                bank_info.append([x.strip() for x in line.split(',')[1:]])
            elif line != '' and not line.startswith('#'):
                key, value = line.split('=')
                user_dict[key.strip()] = value.strip()

    # User Specific Variables
    if user_dict['payroll'].lower() == 'on':
        payroll = True
        pay_period_length = int(user_dict['pay_period_length'])  # days in each pay period
        start_date = date(*[int(x) for x in user_dict['start_date'].split('-')])  # first pay period with current employer
    else:
        payroll = False

    payroll_filename = ''                                       #   user_dict['payroll_filename']   $$$ Not in use, need to enter a get key otherwise default value
    archive_filename = user_dict['archive_filename']
    archive_path = user_dict['archive_path']

    download_path = user_dict['download_path']
    archive_filepath = os.path.join(archive_path, archive_filename)

    label = datetime.now().strftime('%Y_%m_%d')
    archive_title, archive_ext = archive_filename.split('.')
    backup_filename = archive_title + '_' + label + '.' + archive_ext
    backup_filepath = os.path.join(archive_path, backup_filename)
    backups_to_keep = int(user_dict['backups_to_keep'])

    all_banks = list()
    for bank in bank_info:
        extraction_fn = globals()[bank[2]]
        all_banks.append(Bank(bank[0], bank[1], extraction_fn, int(bank[3]), bank[4]))

    # User Independent Variables
    today = datetime.now().strftime('%m/%d/%Y')
    field_names = ['Date', 'Description', 'Amount', 'Category', 'Subcategory', 'Notes', 'Tag', 'Status',
                   'Payperiod', 'Account', 'Added', 'Datenum']
    os.chdir(archive_path)
    # Prepare existing records and create backup copy
    shutil.copyfile(archive_filepath, backup_filepath)
    archive_cleanup(archive_filename, backups_to_keep)
    # Prepare downloaded bank data files and identify most recent
    available_downloads = find_downloads(all_banks, download_path)
    # Import existing records
    if archive_ext == 'csv':
        existing_trans = unpack_transactions(archive_filename)
    elif archive_ext == 'xlsx':
        existing_trans = unpack_transactions(archive_filename, user_dict['archive_sheet'], origin=(4, 2))
    # -- End Setup ------------------------------------------------------

    # -- Update Payroll Deductions --------------------------------------
    new_payroll_trans = list()
    if payroll:
        # Import set of payroll deduction records
        if archive_ext == 'csv':
            payrollset_trans = unpack_transactions(payroll_filename)
        elif archive_ext == 'xlsx':
            payrollset_trans = unpack_transactions(archive_filename, user_dict['payroll_sheet'], origin=(2, 2))
        print('>>> Creating new records for payroll...')
        # Create list of all pay periods from existing records
        included_periods = {record.payp for record in existing_trans if record.acnt == 'Payroll Service'}
        # Calculate all pay periods from first paycheck through today
        expected_periods = periodic_dates(start_date, pay_period_length)
        # Create a list of pay periods that have not yet been added in existing records
        absent_periods = expected_periods - included_periods
        # Copy the set of payroll deduction records for each new period
        new_payroll_trans = list()
        for period in absent_periods:
            batch_setattr(payrollset_trans, 'pubdate', period)
            new_payroll_trans.extend(deepcopy(payrollset_trans))
    # -- End Update Payroll Deductions ----------------------------------

    # -- Import Downloads -----------------------------------------------
    print('>>> Creating new records from downloads...')
    # Iterate over bank data and compile all download records
    #     repeat or also collect func/account values
    downloaded_trans = list()
    downloaded_values = list()
    for bank in available_downloads:
        unpack_data = bank.extraction
        trans, values = unpack_data(bank)
        downloaded_trans.extend(trans)
        downloaded_values.extend(values)
    # Iterate over compiled download records and check in existing
    #     build a list of the download records absent from existing
    new_downloaded_trans = find_missing(downloaded_trans, existing_trans)
    # -- End Import Downloads -------------------------------------------

    # -- Update Market Valuation ----------------------------------------
    print('>>> Creating new records for changes in valuation...')
    # Iterate over account/fund values collected
    #     compute current fund value from existing and new downloaded records
    #     Compile all valuation records when 'download balance' - 'current fund value' is nonzero
    new_valuation_trans = list()
    for present_value in downloaded_values:
        if present_value.tag:
            balance_existing = sum([x.amnt for x in existing_trans if x.tag == present_value.tag])
            balance_new = sum([x.amnt for x in new_downloaded_trans if x.tag == present_value.tag])
            balance_adjustment = round(present_value.amnt-(balance_existing + balance_new), 2)
            if balance_adjustment != float(0):
                # compile all valuation records
                present_value.amnt = balance_adjustment
                new_valuation_trans.append(present_value)
        else:
            pass
            # Do some error checking on account balances that do not change in value like Widget
    # -- End Update Market Valuation ------------------------------------

    # -- New Record Automatic Categorization -----------------------
    print('>>> Running categorization rules and compiling all records...')
    # ***Not yet implemented
    # -- End New Record Automatic Categorization -------------------

    # -- Save All Records to Disk ----------------------------------
    print('>>> Writing all records back to disk...')
    # Mark all new records with today's date
    new_trans = new_payroll_trans + new_downloaded_trans + new_valuation_trans
    batch_setattr(new_trans, 'added', today)
    # Compile all new and preexisting records into a single statement
    all_trans = new_trans + existing_trans
    # Sort by date, write out all transactions to csv
    rows_to_add = len(new_trans)
    print(info + '{} new transactions will be added to {}'.format(rows_to_add, archive_filename))
    if archive_ext == 'csv':
        all_trans = sorted(all_trans, key=lambda item: item.pubdate, reverse=True)
        write_to_csv(archive_filepath, all_trans, field_names)
        os.startfile(archive_filepath)
    elif archive_ext == 'xlsx':
        # Write additions directly to spreadsheet
        # Parameters used while writing to Excel
        #   archive_filepath, archive_sheet, new_trans
        if rows_to_add:
            xl_app = win32com.client.Dispatch('Excel.Application')
            xl_app.Workbooks.Open(archive_filepath)
            xl_sheet = xl_app.Sheets(user_dict['archive_sheet'])
            entry_row = 11
            last_row = entry_row + rows_to_add
            insert_range = str(entry_row)+':'+str(last_row-1)
            xl_sheet.Rows(insert_range).Insert()
            for row, record in enumerate(new_trans):
                row += entry_row
                for col, element in enumerate(record.spoon_feed()):
                    col += 2
                    xl_sheet.Cells(row, col).Formula = str(element)
                stdout.write('\r'+info+'Adding transaction {} of {}'.format(row+1 - entry_row, rows_to_add))
            print('')
            xl_app.Visible = True
    # -- End Save All Records to Disk ------------------------------
    os.chdir(existing_dir)
    print('>>> Processing Complete')

if __name__ == "__main__":
    print('Called from [if] of "updatefinances"')
    main([])

