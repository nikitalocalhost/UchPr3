import PySimpleGUI as sg
from os import walk, path
from xmlmd.template import template
from xmlmd.main import read
from xmlmd.parse import get_teacher_info, merge_teacher_info, get_year, sort_groups, prepare_groups


def get_all_files(dir):
    f = []
    for (dirpath, dirnames, filenames) in walk(dir):
        fn = []
        for file in filenames:
            fn.append(path.join(dir, file))
        f.extend(fn)
        break
    return f


def parse_all_files(files):
    output = {}
    for file in files:
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
    for i in output:
        sorted = sort_groups(output[i])
        # print(sorted)
        year = 2021
        prepared = prepare_groups(sorted, year)
        print(len(prepared) == 0)
        if len(prepared) > 0:
            file = template(i, prepared, year)
            [familia, name, father] = i.split(' ')
            file.save(path.join('o', '%s_%d.xls' % (familia, year)))
    # print(sort_groups(output))

all_files = get_all_files('input')
print(all_files)
parse_all_files(all_files)



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
