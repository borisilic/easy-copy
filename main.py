import argparse
import workbook_helper
import sys

parser = argparse.ArgumentParser()
parser.add_argument('--file', help='The name of the file you wish to work with')
args = parser.parse_args()


if args.file:
    wb = workbook_helper.open_workbook(args.file)
else:
    wb = workbook_helper.open_workbook('default.xlsx')

if wb is None:
    print('Filename is incorrect or file doesn\'t exist. Exiting program.')
    sys.exit()

response = 0
while response != str(3):
    print('Select which sheet you would like to work on: ')
    print('1: Product Master'.rjust(20, " "))
    print('2: BGR'.rjust(9, " "))
    print('3: Exit'.rjust(10, " "))
    response = input()
