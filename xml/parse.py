import re
from typing import Any
from .main import read, find_data_in_col
from xlrd.book import Book
from xlrd.sheet import Sheet


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
            'name': row[2],
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
    k1 = list(gr['semesters'])[0]
    v1 = gr['year'] * 10 + k1 * 4.99
    s1 = gr['name']
    s2 = gr['group']

    return (v1, s1, s2)

def sort_groups(gr: list[dict[str, Any]]):
    return gr.sort(key=_sort_groups)
