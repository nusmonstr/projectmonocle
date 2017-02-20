'''
This is a placeholder for future functions that help
 add manual transactions, split existing transactions, and otherwise modify records

'''

import os
import win32com.client
from sys import stdout
from datetime import date, datetime
from financialelements import Transaction

def alt_input(user_prompt):
    raw_input = input(user_prompt)
    user_input = raw_input.lower().strip()
    if user_input == '#more':
        return False, ''
    else:
        return True, user_input


def main():
    # -- Setup ----------------------------------------------------------
    print('>>> Gathering downloads, payroll definition, and existing records...')
    # Variable declaration
    info = '    [!] '
    question = '[?] '
    response = '\t[select]: '
    # Import User Configuration
    print(question + 'Which user profile would you like to use?')
    user_profiles = [file for file in os.listdir(os.getcwd()) if file.endswith('ini')]
    for i, profile in enumerate(user_profiles):
        print('\t{} {}'.format(i, profile))
    selection = int(input(response))
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

    payroll_filename = ''  # user_dict['payroll_filename']   $$$ Not in use, need to enter a get key otherwise default value
    archive_filename = user_dict['archive_filename']
    archive_path = user_dict['archive_path']

    download_path = user_dict['download_path']
    archive_filepath = os.path.join(archive_path, archive_filename)

    label = datetime.now().strftime('%Y_%m_%d')
    archive_title, archive_ext = archive_filename.split('.')
    backup_filename = archive_title + '_' + label + '.' + archive_ext
    backup_filepath = os.path.join(archive_path, backup_filename)
    backups_to_keep = int(user_dict['backups_to_keep'])

    # User Independent Variables
    today = datetime.now().strftime('%m/%d/%Y')
    new_trans = []
    while input('Enter for next record, [w] to write them to archive').lower().strip() != 'w':
        print('When entering transaction details, you can always enter #more to enter less generic info')
        day = alt_input('Transaction day of the month: ')
        if day[0]:
            pubdate = '{}/{}/{}'.format(datetime.now().month, ('0'+day[1])[-2:], datetime.now().year)  # expecting string format '%m/%d/%Y'
        else:
            pubdate = input('Enter date in mm/dd/yyyy format: ').lower().strip()
        desc = input('Enter a description for this transaction: ').strip()
        amnt = float(input('Enter the amount of this transaction: ').strip())
        cat, subcat = [x.strip() for x in input('Enter a category and subcategory for this transaction (delimit by |): ').split('|')]
        user_input = input('Would you like to use the description for the note? [y]').lower().strip()
        if user_input == 'y':
            memo = desc
        else:
            memo = input('Enter a note for this transaction: ').strip()
        record = Transaction(pubdate, desc, amnt, cat, subcat, notes=memo, status='Liquid', acnt='Cash', added=today)
        print(record)
        user_input = input('Confirm the transaction above by pressing [y]').lower().strip()
        if user_input == 'y':
            new_trans.append(record)

    rows_to_add = len(new_trans)


    if rows_to_add:
        xl_app = win32com.client.Dispatch('Excel.Application')
        xl_app.Workbooks.Open(archive_filepath)
        xl_sheet = xl_app.Sheets(user_dict['archive_sheet'])
        entry_row = 11
        last_row = entry_row + rows_to_add
        insert_range = str(entry_row ) + ':' +str( last_row -1)
        xl_sheet.Rows(insert_range).Insert()
        for row, record in enumerate(new_trans):
            row += entry_row
            for col, element in enumerate(record.spoon_feed()):
                col += 2
                xl_sheet.Cells(row, col).Formula = str(element)
            stdout.write('\r' + info +'Adding transaction {} of {}'.format( row +1, rows_to_add))
        print('')
        xl_app.Visible = True

main()