import datetime
import xlwt
from xlwt.Style import easyxf

now = datetime.datetime.now()
year = now.year

styles = {
    "text": easyxf('font: name Times New Roman; align: wrap on'),
    "header": easyxf('font: name Times New Roman, bold on; align: wrap on, vert centre, horiz center')
}

def template(starting_year: int = year):

    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet('Лист1')

    ws.write_merge(1, 1, 0, 17, 'ПЕДАГОГИЧЕСКАЯ НАГРУЗКА на  %d / %d учебный год' % (starting_year, starting_year + 1), styles["header"])

    return wb

wb = template()
wb.save('output/1.xls')