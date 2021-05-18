import os
import re
from typing import Pattern
import xlrd
from xlrd.book import Book
from xlrd.sheet import Sheet


def read(path: os.PathLike) -> Book:
    extension: str = os.path.splitext(path)[1]
    if extension == '.xls':
        wb = xlrd.open_workbook(path)
    # elif extension == '.xlsx':
    #     wb = oxl.load_workbook(path)
    else:
        raise Exception("File needs to be Excel")

    return wb


def find_data_in_col(sheet: Sheet, col: int, pattern: Pattern) -> dict[int, str]:
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


# wb = read(os.path.join(
#     'input', 'Uchebny_plan_09_02_03_OChNAYa_FORMA_9_klass_2020.xls'))
# ws = wb.sheet_by_index(2)
# pattern = re.compile("[a-zA-Zа-яА-Я]+ [a-zA-Zа-яА-Я]+")
# data = find_data_in_col(ws, 105, pattern)


# print(data)
