import datetime
from typing import Any
import xlwt
from xlwt.Style import easyxf

now = datetime.datetime.now()

styles = {
    "text": easyxf('borders: top thin, bottom thin, left thin, right thin; font: name Times New Roman; align: wrap on'),
    "text_cr": easyxf('borders: top thin, bottom thin, left thin, right thin; font: name Times New Roman; align: wrap on, horiz center'),
    "header": easyxf('borders: top thin, bottom thin, left thin, right thin; font: name Times New Roman, bold on; align: wrap on, vert centre, horiz center'),
    "header_wb": easyxf('font: name Times New Roman, bold on; align: wrap on, vert centre, horiz center'),
    "header_vt": easyxf('borders: top thin, bottom thin, left thin, right thin; font: name Times New Roman, bold on; align: wrap on, vert centre, horiz center, rotation 90')
}

magic_value_col = int(256 / 7)
magic_value_row = int(20 * 72 / 96 * 0.8)


def set_col(ws, n: int, w: int):
    ws.col(n).width = w * magic_value_col


def set_row(ws, n: int, w: int):
    ws.row(n).set_style(easyxf('font: height %d' % (w * magic_value_row)))


def template(fio: str, rows: list[dict[str, Any]], year: int = now.year):

    def write(ws, row: int, col: int, value, style, dv=0):
        if value == dv:
            ws.write(row, col, "", style)
        else:
            ws.write(row, col, value, style)

    def insert_subject(ws, row: int, n: int, subject: str, group: str, group_col: int, sems: tuple[list[int], list[int]], dopr: int, vkr: int, gek: int):
        ws.write(row, 0, n, styles['text_cr'])
        ws.write(row, 1, subject, styles['text'])
        ws.write(row, 2, group, styles['text_cr'])
        ws.write(row, 3, group_col, styles['text_cr'])
        i2 = 0
        for sem in sems:
            i1 = 0
            for v in sem:
                write(ws, row, 4 + i1 + i2 * 5, v, styles['text_cr'])
                i1 += 1
            i2 += 1
        write(ws, row, 14, dopr, styles['text_cr'])
        write(ws, row, 15,  vkr, styles['text_cr'])
        write(ws, row, 16,  gek, styles['text_cr'])
        write(ws, row, 17,  xlwt.Formula('SUM(E%d:Q%d)' % (row+1, row+1)), styles['text_cr'])

    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet('Лист1')

    set_col(ws, 0, 33)
    set_col(ws, 1, 350)
    set_col(ws, 4, 30)
    set_col(ws, 5, 30)
    set_col(ws, 6, 30)
    set_col(ws, 7, 30)
    set_col(ws, 8, 30)
    set_col(ws, 9, 30)
    set_col(ws, 10, 30)
    set_col(ws, 11, 30)
    set_col(ws, 12, 30)
    set_col(ws, 13, 30)
    set_col(ws, 14, 40)
    set_col(ws, 15, 40)
    set_col(ws, 16, 40)
    set_col(ws, 17, 50)
    set_row(ws, 4, 240)

    ws.write_merge(1, 1, 0, 17, 'ПЕДАГОГИЧЕСКАЯ НАГРУЗКА на  %d / %d учебный год' %
                   (year, year + 1), styles["header_wb"])
    ws.write_merge(2, 2, 0, 17, 'Преподаватль %s' % (fio), styles["header_wb"])
    ws.write_merge(3, 4, 0, 0, '№ п/п', styles["header"])
    ws.write_merge(3, 4, 1, 1, 'Предмет, дисциплина, МДК', styles["header"])
    ws.write_merge(3, 4, 2, 2, 'Группа', styles["header"])
    ws.write_merge(3, 4, 3, 3, 'Кол-во обучающихся в группе', styles["header"])

    ws.write_merge(3, 3, 4, 8, '1 семестр', styles['header'])
    ws.write(4, 4, 'Аудиторные занятия', styles['header_vt'])
    ws.write(4, 5, 'Консультации', styles['header_vt'])
    ws.write(4, 6, 'Практика', styles['header_vt'])
    ws.write(4, 7, 'Прием курсовых работ(проектов)', styles['header_vt'])
    ws.write(4, 8, 'Промежуточная аттестация', styles['header_vt'])

    ws.write_merge(3, 3, 9, 13, '2 семестр', styles['header'])
    ws.write(4, 9, 'Аудиторные занятия', styles['header_vt'])
    ws.write(4, 10, 'Консультации', styles['header_vt'])
    ws.write(4, 11, 'Практика', styles['header_vt'])
    ws.write(4, 12, 'Прием курсовых работ(проектов)', styles['header_vt'])
    ws.write(4, 13, 'Промежуточная аттестация', styles['header_vt'])

    ws.write_merge(
        3, 4, 14, 14, 'Домашняя контрольная работа (заочн.)', styles['header_vt'])
    ws.write_merge(3, 4, 15, 15, 'ВРК', styles['header'])
    ws.write_merge(3, 4, 16, 16, 'ГЭК', styles['header'])
    ws.write_merge(3, 4, 17, 17, 'ИТОГО', styles['header_vt'])

    n = 5

    for row in rows:
        insert_subject(ws, n, 6 - n, row['name'],
                       row['group'], row['group_col'], row['sems'], row['dopr'], row['vkr'], row['gek'])
        n += 1
    n += 1
    write(ws, n, 17, xlwt.Formula('SUM(R6:R%d)' % (n+1)), styles['text_cr'])
    n += 1
    write(ws, n, 1, "ИТОГО педагогическая нагрузка на год: ", styles['text'])
    write(ws, n, 2, xlwt.Formula('R%d' % (n)), styles['text_cr'])
    # write(ws, n, 2, "asd", styles['text_cr'])
    n += 1
    write(ws, n, 1, "Заместитель директора", styles['text'])
    n += 1
    write(ws, n, 1, "по УМР:                 _______________Н.Ю. Таратынова",
          styles['text'])
    n += 1
    [familia, name, father] = fio.split(" ")
    write(ws, n, 1, "Преподаватель:     _______________%s.%s.%s" %
          (name[0], father[0], familia), styles['text'])

    return wb


# wb = template("Новиков Арнольд Сергеевич", [{'name': 'МДК.01.02 Прикладное программирование', 'group': '336ПО', 'group_col': 13, 'sems': (
#     [96, 8, 0, 0, 0], [0, 0, 0, 0, 0]), 'dopr': 0, 'vkr': 0, 'gek': 0}])
# wb.save('output/1.xls')
