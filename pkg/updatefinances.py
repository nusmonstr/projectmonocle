'''
Created on Jul 29, 2016
Modified on Apr 23, 2018
@author: Eric Champe
'''
from sys import stdout
from .subpkg.ericsdata_functions import *
from .subpkg.financialelements import unpack_transactions, batch_setattr, Bank, find_missing, periodic_dates, find_downloads, archive_cleanup, midlast_month_weekdays
from datetime import datetime
from copy import deepcopy
import win32com.client
from shutil import copyfile
from os import path, chdir


def explicit_use():   # How do I tell pycharm that the functions are being used, I call the bank functions programmatically with extraction_fn
    import_citi()


def main():
    #####################################################################
    # -- Fetch Online Bank Data -----------------------------------------
    # print('>>> Fetching bank data from the web...')
    # ***Not yet implemented
    # -- End Fetch Online Bank Data -------------------------------------

    #####################################################################
    # -- Setup ----------------------------------------------------------
    print('>>> Gathering downloads, payroll definition, and existing records...')
    # Variable declaration
    today = datetime.now()
    script_dir = path.dirname(path.abspath(__file__))
    config_filename = 'config_eric.ini'
    bank_info = []
    with open(path.join(script_dir, config_filename), 'r') as config:
        lines = [x.strip() for x in config.readlines()]
        user_dict = dict()
        for line in lines:
            if line.startswith('<Bank>'):
                bank_info.append([x.strip() for x in line.split(',')[1:]])
            elif line != '' and not line.startswith('#'):
                key, value = line.split('=')
                user_dict[key.strip()] = value.strip()
    user_dict['pay_type'] = userpayp    # declared in ericsdata_functions
    archive_filepath = path.join(user_dict['archive_path'], user_dict['archive_filename'])
    backup_filepath = path.join(user_dict['archive_path'], user_dict['archive_filename'].replace('.xlsx','_' + today.strftime('%Y_%m_%d') + '.xlsx'))
    all_banks = list()
    for bank in bank_info:
        extraction_fn = globals()[bank[2]]
        all_banks.append(Bank(bank[0], bank[1], extraction_fn, int(bank[3]), bank[4]))
    chdir(user_dict['archive_path'])
    # Prepare existing records and create backup copy
    copyfile(archive_filepath, backup_filepath)
    archive_cleanup(user_dict['archive_filename'], int(user_dict['backups_to_keep']))
    # Prepare downloaded bank data files and identify most recent
    available_downloads = find_downloads(all_banks, user_dict['download_path'])
    # Import existing records
    existing_trans = unpack_transactions(user_dict['archive_filename'], user_dict['archive_sheet'], origin=(4, 2))
    # -- End Setup ------------------------------------------------------

    #####################################################################
    # -- Update Payroll Deductions --------------------------------------
    print('>>> Creating new records for payroll...')
    # Import set of payroll deduction records
    payrollset_trans = unpack_transactions(user_dict['archive_filename'], user_dict['payroll_sheet'], origin=(2, 2))
    # Create list of all pay periods from existing records
    included_periods = {record.payp for record in existing_trans if record.acnt == 'Payroll Service' and record.paypstyle == user_dict['pay_type']}
    # Calculate all pay periods from first paycheck through today
    if user_dict['pay_type'] == 'Semimonthly':
        expected_periods = midlast_month_weekdays(user_dict['start_date'])
    if user_dict['pay_type'].startswith('Biweekly'):
        expected_periods = periodic_dates(start_date, pay_period_length=14)
    # Create a list of pay periods that have not yet been added in existing records
    absent_periods = expected_periods - included_periods
    # Copy the set of payroll deduction records for each new period
    new_payroll_trans = list()
    for period in absent_periods:
        batch_setattr(payrollset_trans, 'pubdate', period)
        new_payroll_trans.extend(deepcopy(payrollset_trans))
    # -- End Update Payroll Deductions ----------------------------------

    #####################################################################
    # -- Import Downloads -----------------------------------------------
    print('>>> Creating new records from downloads...')
    # Iterate over bank data and compile all download records
    downloaded_trans = list()
    downloaded_values = list()
    for bank in available_downloads:
        unpack_data = bank.extraction
        trans, values = unpack_data(bank)
        downloaded_trans.extend(trans)
        downloaded_values.extend(values)
    # Check download records against existing and keep only new transactions
    new_downloaded_trans = find_missing(downloaded_trans, existing_trans)
    # -- End Import Downloads -------------------------------------------

    #####################################################################
    # -- Update Market Valuation ----------------------------------------
    print('>>> Creating new records for changes in valuation...')
    # Iterate over account/fund values collected
    #     compute current fund value from existing and new downloaded records
    #     Compile all valuation records when 'download balance' - 'current fund value' is nonzero
    new_valuation_trans = list()
    current_trans = existing_trans + new_downloaded_trans
    for present_value in downloaded_values:
        balance_current = sum([x.amnt for x in current_trans if x.fund == present_value.fund and x.acnt == present_value.acnt])
        balance_adjustment = round(present_value.amnt-balance_current, 2)
        #print('____{}\n'.format(present_value.desc + ' ' + present_value.acnt), 'Bal file:', balance_current, '\nBal Download:', present_value.amnt, '\nBal Adjustment:', balance_adjustment, '\nAdjustment Needed:', balance_adjustment != float(0))
        if balance_adjustment != float(0):
            # compile all valuation records
            present_value.amnt = balance_adjustment
            new_valuation_trans.append(present_value)
    # -- End Update Market Valuation ------------------------------------

    #####################################################################
    # -- New Record Automatic Categorization -----------------------
    #print('>>> Running categorization rules and compiling all records...')
    # ***Not yet implemented
    # -- End New Record Automatic Categorization -------------------

    #####################################################################
    # -- Save All Records to Disk ----------------------------------
    print('>>> Writing all records back to disk...')
    # Compile all new records and populate "added" with today's date
    new_trans = new_payroll_trans + new_downloaded_trans + new_valuation_trans
    batch_setattr(new_trans, 'added', today.strftime('%m/%d/%Y'))
    # Write additions directly to spreadsheet
    print('    [!] ' + '{} new transactions will be added to {}'.format(len(new_trans), user_dict['archive_filename']))
    if new_trans:
        xl_app = win32com.client.dynamic.Dispatch('Excel.Application')
        xl_app.Workbooks.Open(archive_filepath)
        xl_sheet = xl_app.Sheets(user_dict['archive_sheet'])
        xl_app.Visible = False
        for i, record in enumerate(new_trans):
            xl_sheet.Rows('11:11').Insert()
            xl_sheet.Range('B11:P11').value = record.spoon_feed()    # record is 15 elements long
            stdout.write('\r    [!] ' + 'New Transaction #{} is being written'.format(i))
        xl_app.Visible = True
    else:
        print('    [!] '+'No new transactions were found.')
    print('>>> Processing Complete')
    # -- End Save All Records to Disk ------------------------------


if __name__ == "__main__":
    print('Call from finpy instead.')

