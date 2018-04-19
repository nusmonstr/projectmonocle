#from nose import tests
import os
import sys
sys.path.insert(0, os.path.abspath('..'))
from pkg import updatefinances
#from ..shwrap import module1, module2


def test1(args):
    print('Test module running on 1')
    updatefinances.main(args)

def test2():
    print('Test module running on 2')
    pass    #module2.main()

if __name__ == "__main__":
    print('Called from [if] of "tests/run"')
    test1(['empty test run'])

'''
To change current working dir to the one containing your script you can use:
import os
os.chdir(os.path.dirname(__file__))
print(os.getcwd())
'''