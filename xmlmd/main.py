import os
import re
from typing import Pattern
import xlrd
from xlrd.book import Book
from xlrd.sheet import Sheet

# import openpyxl as oxl
# from openpyxl import Workbook


def read_xls(path: os.PathLike) -> Book:
    extension: str = os.path.splitext(path)[1]
    if extension == '.xls':
        wb = xlrd.open_workbook(path)
    else:
        raise Exception("File needs to be Excel")

    return wb


# def read_xlsx(path: os.PathLike) -> Workbook:
#     extension: str = os.path.splitext(path)[1]
#     if (extension) == '.xlsx':
#         wb = oxl.load_workbook(path)
#     else:
#         raise Exception("File needs to be Excel")

#     return wb

def find_data_in_col_xls(sheet: Sheet, col: int, pattern: Pattern) -> dict[int, str]:
    data: dict[int, str] = {}
    for i in range(sheet.ncols):
        value = sheet.cell_value(i, col)
        if value:
            try:
                if (pattern.search(value)):
                    data[i] = value
            except:
                pass
    return data

# def find_data_in_col_xlsx(sheet, col: int, pattern: Pattern) -> dict[int, str]:
#     data: dict[int, str] = {}
#     for i in range(sheet.ncols):
#         value = sheet.cell_value(i, col)
#         if value:
#             try:
#                 if (pattern.search(value)):
#                     data[i] = value
#             except:
#                 pass
#     return data


# wb = read_xls(os.path.join(
#     'input', 'Uchebny_plan_09_02_03_OChNAYa_FORMA_9_klass_2020.xls'))
# ws = wb.sheet_by_index(2)
# pattern = re.compile("[a-zA-Zа-яА-Я]+ [a-zA-Zа-яА-Я]+")
# data = find_data_in_col_xls(ws, 105, pattern)


# print(data)
