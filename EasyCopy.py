import argparse
import pyperclip
from pynput import keyboard
import string
import workbook_helper
import sys

parser = argparse.ArgumentParser()
parser.add_argument('--file', help='The name of the file you wish to work with')
parser.add_argument('--sheet', help='The name of the sheet you wish to open')
args = parser.parse_args()


wb = workbook_helper.get_file(args)

if wb is None:
    print('Could not find file. Exiting program.')
    sys.exit()

if args.sheet:
    sheet = wb[args.sheet]
else:
    sheets = wb.sheetnames
    sheet = wb[sheets[0]]


rowNumber = range(1, sheet.max_row + 1)
headings = []
products = []
lineNumber = 0

for cell in sheet[1]:
    headings.append(cell.value)


def to_uppercase(index):
    return string.ascii_letters[index].upper()


for row in rowNumber:
    line = dict()
    for heading in headings:
        cell_value = sheet[to_uppercase(headings.index(heading)) + str(row)].value
        line[heading] = cell_value
    products.append(line)


print(products)


def on_press(key):
    global lineNumber
    if key == keyboard.Key.shift_r:
        copy_code()
    if key == keyboard.Key.ctrl_r:
        copy_description()
    if key == keyboard.Key.left:
        get_previous_line()
    if key == keyboard.Key.right:
        get_next_line()


def copy_code():
    reece_code = products[lineNumber]['Reece Code']
    pyperclip.copy(reece_code)
    print('Copied: ' + str(reece_code))


def copy_description():
    description = products[lineNumber]['Description']
    pyperclip.copy(description)
    print('Copied: ' + str(description))


def get_previous_line():
    global lineNumber
    try:
        lineNumber = lineNumber - 1
        line_info = products[lineNumber]
        display_line(lineNumber, line_info)
    except IndexError:
        lineNumber = lineNumber + 1
        print('Reached end of list.')


def get_next_line():
    global lineNumber
    try:
        lineNumber = lineNumber + 1
        line_info = products[lineNumber]
        display_line(lineNumber, line_info)
    except IndexError:
        lineNumber = lineNumber - 1
        print('Reached end of list.')


def display_line(number, info):
    print()
    print('================')
    print('Current line:')
    print(str(number) + ":", str(info))
    print('================')


with keyboard.Listener(
        on_press=on_press) as listener:
    listener.join()
