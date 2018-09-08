import pyperclip
from pynput import keyboard
import workbook_helper
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('--file', help='The name of the file you wish to work with')
parser.add_argument('--sheet', help='The name of the sheet you wish to open')
args = parser.parse_args()

wb = workbook_helper.get_file(args)

sheet = wb['BGR']

rowNumber = range(1, sheet.max_row + 1)
headings = workbook_helper.set_headings(sheet)
products = workbook_helper.set_products(headings, rowNumber, sheet)
lineNumber = 0


def on_press(key):
    global lineNumber
    if key == keyboard.Key.shift_r:
        copy_part_number()
    if key == keyboard.Key.ctrl_r:
        copy_invoice_cost()
    if key == keyboard.Key.end:
        copy_rebate()
    if key == keyboard.Key.left:
        lineNumber = workbook_helper.get_previous_line(lineNumber, products)
    if key == keyboard.Key.right:
        lineNumber = workbook_helper.get_next_line(lineNumber, products)


def copy_part_number():
    part_number = products[lineNumber][headings[3]]
    pyperclip.copy(part_number)
    print('Copied: ' + str(part_number))


def copy_invoice_cost():
    try:
        invoice_cost = round(float(products[lineNumber][headings[4]]), 2)
        pyperclip.copy(invoice_cost)
        print('Copied: ' + str(invoice_cost))
    except ValueError:
        print('Cannot copy, this is not a number.')


def copy_rebate():
    rebate = products[lineNumber][headings[5]]
    pyperclip.copy(rebate)
    print('Copied: ' + str(rebate))


# Collect events until released
with keyboard.Listener(
        on_press=on_press) as listener:
    listener.join()
