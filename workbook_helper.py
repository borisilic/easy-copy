import openpyxl


def open_workbook(name='workbook.xlsx'):
    return openpyxl.load_workbook(name)
