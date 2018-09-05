import openpyxl


def open_workbook(name):
    try:
        return openpyxl.load_workbook(name)
    except OSError:
        return None
