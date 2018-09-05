import openpyxl
import pyperclip

import workbook_helper
from pynput import keyboard, mouse



workbookName = 'asdf'
wb = workbook_helper.open_workbook(workbookName)

sheet1 = wb['Product Master']

rowNumber = range(1, sheet1.max_row + 1)

products = []

for row in rowNumber:
    productNumbersColumn = 'A' + str(row)
    descriptionColumn = 'B' + str(row)
    products.append({
        'reeceCode': sheet1[productNumbersColumn].value,
        'description': sheet1[descriptionColumn].value
    })


print(products)
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
    reece_code = products[lineNumber]['reeceCode']
    pyperclip.copy(reece_code)
    print('Copied: ' + str(reece_code))


def copy_description():
    description = products[lineNumber]['description']
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


# Collect events until released
with keyboard.Listener(
        on_press=on_press) as listener:
    listener.join()
