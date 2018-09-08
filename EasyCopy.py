import argparse
import pyperclip
from pynput import keyboard
import workbook_helper

parser = argparse.ArgumentParser()
parser.add_argument('--file', help='The name of the file you wish to work with')
parser.add_argument('--sheet', help='The number of the sheet you wish to open (1 or 2 or 3 or..)')
parser.add_argument('--columns', help='The columns you want to be able to copy with keyboard')
args = parser.parse_args()

wb = workbook_helper.get_file(args)
sheet = workbook_helper.get_sheet(args, wb)
rowNumber = range(1, sheet.max_row + 1)
headings = workbook_helper.set_headings(sheet)
products = workbook_helper.set_products(headings, rowNumber, sheet)
headings_to_copy = workbook_helper.set_columns_to_copy(args, headings)
lineNumber = 0


def on_press(key):
    global lineNumber
    if key == keyboard.Key.shift_r:
        copy(0)
    if key == keyboard.Key.ctrl_r or key == keyboard.Key.cmd_r:
        copy(1)
    if key == keyboard.Key.end:
        copy(2)
    if key == keyboard.Key.left:
        lineNumber = workbook_helper.get_previous_line(lineNumber, products)
    if key == keyboard.Key.right:
        lineNumber = workbook_helper.get_next_line(lineNumber, products)


def copy(element_requested):
    element = products[lineNumber][headings_to_copy[element_requested]]
    pyperclip.copy(element)
    print('Copied: ' + str(element))


with keyboard.Listener(
        on_press=on_press) as listener:
    listener.join()
