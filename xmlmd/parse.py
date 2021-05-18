import re
import datetime
from typing import Any
from .main import read, find_data_in_col
from xlrd.book import Book, expand_cell_address
from xlrd.sheet import Sheet

now = datetime.datetime.now()

name_pattern = re.compile("[a-zA-Zа-яА-Я]+ [a-zA-Zа-яА-Я]+")


def get_year(wb: Book):
    ws = wb.sheet_by_index(0)
    year = ws.cell_value(26, 44)
    return int(year)


def get_teacher_info(ws: Sheet, year: int, group: str, group_col: int):
    teachers: dict[str, list] = {}
    sems = 8
    data = False
    while not data:
        try:
            data = find_data_in_col(ws, 105 - ((8 - sems) * 10), name_pattern)
        except:
            sems -= 1

    # try:
    #     data = find_data_in_col(ws, 105, name_pattern)
    # except:
    #     data = find_data_in_col(ws, 105 - 20, name_pattern)
    #     sems -= 2
    for i in data:
        name = data[i]
        if name not in teachers:
            teachers[name] = []

        teacher = teachers[name]
        row = ws.row_values(i)

        is_pr = True if row[7] else False

        semesters = {}

        for i in range(sems):
            x = 20 + i * 10
            sem = row[x:x+10]
            if is_pr:
                hours = sem[3]
                try:
                    if hours:
                        hours_int = int(hours)
                        semesters[i] = hours_int
                    else:
                        pass
                except:
                    pass

            else:
                sem_int = []
                for v in sem:
                    try:
                        if v:
                            v_int = int(v)
                            sem_int.append(v_int)
                        else:
                            sem_int.append(False)
                    except:
                        sem_int.append(False)
                if list(filter(lambda hours: True if hours else False, sem_int)) and sem_int[2] and sem_int[3]:
                    semesters[i] = [sem_int[3], sem_int[2]]
                    # print('%s (%d): %d konsult %d lect' % (name, i+1, sem_int[2], sem_int[3]))

        subject = {
            'group': group,
            'group_col': group_col,
            'year': year,
            'name': row[1] + ' ' + row[2],
            'is_pr': is_pr,
            'semesters': semesters
        }

        teacher.append(subject)

    return teachers


def merge_teacher_info(tl1: dict[str, list], tl2: dict[str, list]):
    list: dict[str, list] = {}
    for l in [tl1, tl2]:
        for t in l:
            if t in list:
                list[t].extend(l[t])
            else:
                list[t] = l[t]
    return list


# def _sort_groups(gr1: dict[str, Any], gr2: dict[str, Any]):
#     semv1: dict[int, list[int]] = gr1['semesters']
#     k1 = semv1.keys[0]
#     v1 = gr1['year'] * 10 + k1 * 4.99

#     semv2: dict[int, list[int]] = gr2['semesters']
#     k2 = semv2.keys[0]
#     v2 = gr2['year'] * 10 + k2 * 4.99

#     if v1 > v2:
#         return 1
#     elif v2 > v1:
#         return -1

#     s1 = gr1['name']
#     s2 = gr2['name']

#     if not s1 > s2:
#         return 1
#     elif not s2 > s1:
#         return -1
#     else:
#         return 0


def _sort_groups(gr: dict[str, Any]):
    semv1: dict[int, list[int]] = gr['semesters']
    try:
        k1 = list(gr['semesters'])[0]
        v1 = gr['year'] * 10 + k1 * 4.99
        s1 = gr['name']
        s2 = gr['group']

        return (v1, s1, s2)
    except:
        return (9999, gr['name'], gr['group'])


def sort_groups(gr: list[dict[str, Any]]):
    gr.sort(key=_sort_groups)
    return gr


def prepare_groups(gr: list[dict[str, Any]], year: int = now.year):
    ngs = []
    for group in gr:
        result1 = [0, 0, 0, 0, 0]
        result2 = [0, 0, 0, 0, 0]
        d_year = year - group['year']
        if d_year < 0:
            pass
        s1 = d_year * 2 + 1
        s2 = d_year * 2 + 2
        # print(d_year, s1, s2)
        if s1 in group['semesters']:
            v1 = group['semesters'][s1]
            if type(v1) == list:
                result1[0] = v1[0]
                result1[1] = v1[1]
            else:
                result1[2] = v1
        if s2 in group['semesters']:
            v2 = group['semesters'][s2]
            if type(v2) == list:
                result2[0] = v2[0]
                result2[1] = v2[1]
            else:
                result2[2] = v2
        sm = 0
        for i in result1:
            sm += i
        for i in result2:
            sm += i
        if sm > 0:
            ng = {
                'name': group['name'],
                'group': group['group'],
                'group_col': group['group_col'],
                'sems': (result1, result2),
                'dopr': 0,
                'vkr': 0,
                'gek': 0
            }
            ngs.append(ng)
    return ngs
