# import PySimpleGUI as sg

# sg.theme('System Default For Real')


# def main():
#     layout = [
#         [sg.FilesBrowse('Выберите файлы...', key='-FILES-', file_types=(("Файлы Excel", "*.xls *xlsx"),))],
#         [sg.B('Сформировать отчетность'), sg.B('Отменить')]
#     ]
#     window = sg.Window('337POk1', layout)

#     while True:
#         event, values = window.read(timeout = 20)
#         print(values['-FILES-'])
#         if values['-FILES-']:
#             window['-FILES-'].update('Выбрано файлов: ' + str(len(values['-FILES-'].split(';'))))
#         else:
#             window['-FILES-'].update('Выберите файлы...')

#         if event == sg.WIN_CLOSED or event == 'Отменить':
#             break
#     window.close()


# if __name__ == '__main__':
#     main()

from os import walk, path
from xml.main import read
from xml.parse import get_teacher_info, merge_teacher_info, get_year


def get_all_files(dir):
    f = []
    for (dirpath, dirnames, filenames) in walk(dir):
        f.extend(filenames)
        break
    return f


def parse_all_files(dir):
    files = get_all_files(dir)
    output = {}
    for e in files:
        file = path.join(dir, e)
        basename = path.basename(file)
        (base, _) = path.splitext(basename)
        [group, group_col] = base.split('_')
        try:
            wb = read(file)
            ws = wb.sheet_by_index(2)
            ti = get_teacher_info(ws, get_year(wb), group, group_col)
            output = merge_teacher_info(output, ti)
        except:
            pass

    print(output)


parse_all_files('input')