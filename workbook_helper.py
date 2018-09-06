import string
import sys
import openpyxl


def get_file(args):
    if args.file:
        return open_workbook(args.file)
    else:
        return open_workbook('default.xlsx')


def open_workbook(name):
    try:
        print('Opening file with name: ' + str(name))
        return openpyxl.load_workbook(name)
    except OSError:
        print('Could not find file. Exiting program.')
        sys.exit()


def set_headings(sheet):
    headings = []
    for cell in sheet[1]:
        headings.append(cell.value)
    return headings


def to_uppercase(index):
    return string.ascii_letters[index].upper()


def set_products(headings, row_number, sheet):
    products = []
    for row in row_number:
        line = dict()
        for heading in headings:
            cell_value = sheet[to_uppercase(headings.index(heading)) + str(row)].value
            line[heading] = cell_value
        products.append(line)
    return products


def get_previous_line(line_number, products):
    try:
        line_number = line_number - 1
        line_info = products[line_number]
        display_line(line_number, line_info)
        return line_number
    except IndexError:
        line_number = line_number + 1
        print('Reached end of list.')
        return line_number


def get_next_line(line_number, products):
    try:
        line_number = line_number + 1
        line_info = products[line_number]
        display_line(line_number, line_info)
        return line_number
    except IndexError:
        line_number = line_number - 1
        print('Reached end of list.')
        return line_number


def display_line(number, info):
    print()
    print('=' * 50)
    print('Current line: ' + str(number + 1))
    # print(str(number) + ":", str(info))
    pretty(info, 1)
    print('=' * 50)


def pretty(d, indent=0):
    for key, value in d.items():
        print('\t' * indent + str(key))
        if isinstance(value, dict):
            pretty(value, indent+1)
        else:
            print('\t' * (indent+1) + str(value))