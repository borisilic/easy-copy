import pyperclip
from pynput import keyboard
import argparse
import sys
import string
import workbook_helper

parser = argparse.ArgumentParser()
parser.add_argument('--file', help='The name of the file you wish to work with')
parser.add_argument('--sheet', help='The name of the sheet you wish to open')
args = parser.parse_args()

if args.file:
    wb = workbook_helper.open_workbook(args.file)
else:
    wb = workbook_helper.open_workbook('default.xlsx')

if wb is None:
    print('Could not find file. Exiting program.')
    sys.exit()

sheet = wb['Product Master']

rowNumber = range(1, sheet.max_row + 1)

headings = []
products = []

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

lineNumber = 0


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
    reece_code = products[lineNumber][headings[0]]
    pyperclip.copy(reece_code)
    print('Copied: ' + str(reece_code))


def copy_description():
    description = products[lineNumber][headings[1]]
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


def pretty(d, indent=0):
    for key, value in d.items():
        print('\t' * indent + str(key))
        if isinstance(value, dict):
            pretty(value, indent+1)
        else:
            print('\t' * (indent+1) + str(value))


def display_line(number, info):
    print()
    print('=' * 50)
    print('Current line: ' + str(number))
    # print(str(number) + ":", str(info))
    pretty(info, 1)
    print('=' * 50)


# Collect events until released
with keyboard.Listener(
        on_press=on_press) as listener:
    listener.join()
