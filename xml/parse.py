import re
from .main import read, find_data_in_col
from xlrd.sheet import Sheet


name_pattern = re.compile("[a-zA-Zа-яА-Я]+ [a-zA-Zа-яА-Я]+")


def get_teacher_info(ws: Sheet, group: str, group_col: int):
    teachers: dict[str, list] = {}
    data = find_data_in_col(ws, 105, name_pattern)
    for i in data:
        name = data[i]
        if name not in teachers:
            teachers[name] = []

        teacher = teachers[name]
        row = ws.row_values(i)

        is_pr = True if row[7] else False

        semesters = {}

        for i in range(8):
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
            'name': row[2],
            'is_pr': is_pr,
            'semesters': semesters
        }

        teacher.append(subject)

    return teachers


# def merge_teacher_info(tl1: dict[str, list], tl2: dict[str, list]):
#     for t1 in tl1:

