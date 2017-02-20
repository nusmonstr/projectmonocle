import sys, os
#print('__file__:', __file__)
#print(__name__)
#print(os.getcwd())
from pkg import updatefinances
from tests import run


def main():
    print('Number of arguments:', len(sys.argv), 'arguments.')
    print('Argument List:', str(sys.argv))


if __name__ == "__main__":
    #print('Called from [if] of "finpy"')
    if len(sys.argv) == 1:
        sys.argv.append('eric')
        sys.argv.append('delete')
        #exit(0)
    if len(sys.argv) > 1 and sys.argv[1] == 'test':
        run.test1(sys.argv)
    else:
        updatefinances.main(sys.argv)
