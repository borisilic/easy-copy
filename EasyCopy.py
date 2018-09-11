import os, sys
import pyperclip
from pynput import keyboard
import workbook_helper


application_path = os.path.dirname(sys.executable)
file = input('Enter file name.\n')
sheet_1 = input('Enter sheet name.\n')
args = input('Enter columns to copy.\n')
wb = workbook_helper.get_file(os.path.join(application_path, file))
sheet = workbook_helper.get_sheet(sheet_1, wb)
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
    try:
        element = products[lineNumber][headings_to_copy[element_requested]]
        pyperclip.copy(element)
        print('Copied: ' + str(element))
    except IndexError:
        print('That key has not been assigned a column to copy')
    

with keyboard.Listener(
        on_press=on_press) as listener:
    listener.join()

input('Press ENTER to exit')
