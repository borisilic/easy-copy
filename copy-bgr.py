import openpyxl
import pyperclip
from pynput import keyboard, mouse
import workbook_helper

wb = openpyxl.load_workbook('Wilson Ads  - Boris - 1203140 to 1203163.xlsx')

sheet2 = wb['BGR']

rowNumber = range(1, sheet2.max_row + 1)
products = []

for row in rowNumber:
    reeceCodeColumn = 'B' + str(row)
    supplierPartColumn = 'D' + str(row)
    invoiceCostColumn = 'E' + str(row)
    products.append({
        'reeceCode': sheet2[reeceCodeColumn].value,
        'supplierPartNumber': sheet2[supplierPartColumn].value,
        'invoiceCost': sheet2[invoiceCostColumn].value
    })


print(products)
lineNumber = 0


def on_press(key):
    global lineNumber
    if key == keyboard.Key.shift_r:
        copy_part_number()
    if key == keyboard.Key.ctrl_r:
        copy_invoice_cost()
    if key == keyboard.Key.left:
        get_previous_line()
    if key == keyboard.Key.right:
        get_next_line()


def copy_part_number():
    part_number = products[lineNumber]['supplierPartNumber']
    pyperclip.copy(part_number)
    print('Copied: ' + str(part_number))


def copy_invoice_cost():
    try:
        invoice_cost = round(float(products[lineNumber]['invoiceCost']), 2)
        pyperclip.copy(invoice_cost)
        print('Copied: ' + str(invoice_cost))
    except ValueError:
        print('Cannot copy, this is not a number.')


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
