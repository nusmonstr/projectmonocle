

def report_on_payperiods(count):
    msg = '{} pay periods will be added in this update.'.format(count)
    msg_one = '{} pay period will be added in this update.'.format(count)
    msg_none = 'No new pay periods will be added in this update.'
    if count == 1: print(msg_one)
    elif count: print(msg)
    else: print(msg_none)


def report_on_existing(count):
    msg = '{} existing records have been imported.'.format(count)
    msg_one = '{} existing record has been imported.'.format(count)
    msg_none = 'No existing records have been imported.'
    if count == 1: print(msg_one)
    elif count: print(msg)
    else: print(msg_none)