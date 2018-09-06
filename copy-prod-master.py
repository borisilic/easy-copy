import pyperclip
from pynput import keyboard
import argparse
import workbook_helper

parser = argparse.ArgumentParser()
parser.add_argument('--file', help='The name of the file you wish to work with')
parser.add_argument('--sheet', help='The name of the sheet you wish to open')
args = parser.parse_args()

wb = workbook_helper.get_file(args)

sheet = wb['Product Master']

rowNumber = range(1, sheet.max_row + 1)
headings = workbook_helper.set_headings(sheet)
products = workbook_helper.set_products(headings, rowNumber, sheet)
lineNumber = 0


def on_press(key):
    global lineNumber
    if key == keyboard.Key.shift_r:
        copy_code()
    if key == keyboard.Key.ctrl_r:
        copy_description()
    if key == keyboard.Key.left:
        lineNumber = workbook_helper.get_previous_line(lineNumber, products)
    if key == keyboard.Key.right:
        lineNumber = workbook_helper.get_next_line(lineNumber, products)


def copy_code():
    reece_code = products[lineNumber][headings[0]]
    pyperclip.copy(reece_code)
    print('Copied: ' + str(reece_code))


def copy_description():
    description = products[lineNumber][headings[1]]
    pyperclip.copy(description)
    print('Copied: ' + str(description))


# Collect events until released
with keyboard.Listener(
        on_press=on_press) as listener:
    listener.join()
