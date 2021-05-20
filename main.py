from os import walk, path
from datetime import datetime
import traceback

import PySimpleGUI as sg

from xmlmd.template import template
from xmlmd.main import read_xls
from xmlmd.parse import get_teacher_info_xls, merge_teacher_info, get_year_xls, sort_groups, prepare_groups

now = datetime.now()

# def get_all_files(dir):
#     f = []
#     for (dirpath, dirnames, filenames) in walk(dir):
#         fn = []
#         for file in filenames:
#             fn.append(path.join(dir, file))
#         f.extend(fn)
#         break
#     return f


def parse_all_files(files, year, of):
    output = {}
    for file in files:
        basename = path.basename(file)
        (base, _) = path.splitext(basename)
        [group, group_col] = base.split('_')
        try:
            wb = read_xls(file)
            ws = wb.sheet_by_index(2)
            ti = get_teacher_info_xls(ws, get_year_xls(wb), group, group_col)
            output = merge_teacher_info(output, ti)
        except:
            pass

    # print(output)
    for i in output:
        sorted = sort_groups(output[i])
        prepared = prepare_groups(sorted, year)
        # print(len(prepared) == 0)
        if len(prepared) > 0:
            file = template(i, prepared, year)
            [familia, name, father] = i.split(' ')
            file.save(path.join(of, '%s_%d.xls' % (familia, year)))
    # print(sort_groups(output))

# all_files = get_all_files('input')
# print(all_files)
# parse_all_files(all_files)

def get_int(s: str) -> int:
    def c(s: str) -> bool:
        return s.isnumeric()

    def g(s: str) -> int:
        return int(s)
    n = 0

    while s:
        if c(s[0]):
            n *= 10
            n += g(s[0])
        s = s[1:]
    return n

def gen(sg, values):
    if not values['-FILES-'] or len(values['-FILES-'].split(';')) == 0:
        sg.PopupError('Файлы не выбраны.', title='Ошибка')
        return
    files = values['-FILES-'].split(';')
    if not values['-YEAR-']:
        sg.PopupError('Год не выбран.', title='Ошибка')
        return
    year = get_int(values['-YEAR-'])
    if not year or year == 0:
        sg.PopupError('Неправильно указан год.', title='Ошибка')
        return
    folder = sg.PopupGetFolder('Выберите папку, куда сохранить')
    print(folder)
    if not folder:
        sg.PopupError('Не выбрана выходная папка.', title='Ошибка')
        return

    try:
        parse_all_files(files, year, folder)
        sg.PopupOK('Успешно')
    except Exception as e:
        tb = traceback.format_exc()
        sg.Print('Ошибка: ', e, tb)


sg.theme('System Default For Real')


def main():
    layout = [
        [sg.FilesBrowse('Выберите файлы...', key='-FILES-',
                        file_types=(("Файлы Excel", "*.xls *xlsx"),))],
        [sg.T('Год: '), sg.I('%d' % now.year, key='-YEAR-',
                             enable_events=True), sg.T(' / %d' % (now.year + 1), key='-YEAR-END-')],
        # [sg.I(''), sg.FolderBrowse('Выберите папку...', key='-FOLDER-')],
        [sg.B('Отменить', key='-CLOSE-'), sg.B('Сформировать', key='-DO-')]
    ]
    window = sg.Window('Тарификация преподавателей', layout)

    while True:
        event, values = window.read(timeout=20)

        if event == sg.WIN_CLOSED or event == '-CLOSE-':
            break

        if values['-FILES-']:
            window['-FILES-'].update('Выбрано файлов: ' +
                                     str(len(values['-FILES-'].split(';'))))
        else:
            window['-FILES-'].update('Выберите файлы...')

        
        if event == '-YEAR-':
            if values['-YEAR-']:
                year = get_int(values['-YEAR-'])
                window['-YEAR-'].update(year)
                window['-YEAR-END-'].update(year + 1)
        if event == '-DO-':
            gen(sg, values)
                

    window.close()


if __name__ == '__main__':
    main()
