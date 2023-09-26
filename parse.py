from copy import copy

import pandas as pd
import numpy as np
import json

python_dict = {
    'name': str,
    'holding_type': str,  # Лекция, сем, лаба
    'course': int,
    'group_name': str,
    'total_time_for_group': int,
    'semester': int,
    'used': False,
    'study_level': str,
    'semester_duration': int,
    'subject_type': str,        # Базовая компонента, Вариативная компонента...
    'department': str,          # Механики и процессов управления
    'credit': int,
    'cipher': str,  # 01.03.02
    'direction': str  # ПМИ
}

list_of_dicts = []

df = pd.read_excel('t1.xls', 'TDSheet')
col_names = [column for column in df][4:]
f_col = [column for column in df][0]
# print(col_names)


def get_dep():
    dep_list = []
    mpu_list = []
    for cell in df['Unnamed: 3'].items():
        if cell[1] is not np.nan:
            dep_list.append(cell)
    for cell in dep_list:
        if cell[1] == 'Механики и процессов управления' or cell[1] == 'Mechanics and Control Processes':
            mpu_list.append(cell)
    return mpu_list


def get_sem(sstr):
    sstr_split = sstr.strip().split()
    return sstr_split[0]


def get_sem_dur(sstr):
    sstr_split = sstr.strip().split()
    try:
        output = sstr_split[-2]
        return output
    except IndexError:
        return 0


def get_course(sstr):
    sstr_split = sstr.strip().split()
    try:
        output = int(sstr_split[-1])
        return output
    except Exception:
        output = int(sstr_split[0][0])
        return output


def get_group_name(sstr):
    sstr_split = sstr.strip().split()
    output = sstr_split[-1]
    return output


def get_study_level(sstr):
    sstr_split = sstr.strip().split()
    output = sstr_split[-1][3]
    if output == 'б':
        return 'Бакалавриат'
    else:
        return 'Магистратура'


def to_full_holding_type(sstr):
    if sstr == 'Пр.' or sstr == 'Sem':
        return 'Семинар'
    if sstr == 'Лек' or sstr == 'Lec':
        return 'Лекция'
    if sstr == 'Лаб' or sstr == 'Lab':
        return 'Лаборатораная работа'


def get_direction(sstr):
    if sstr[0:3] == 'ИУС':
        return 'Управление в технических системах'
    if sstr[0:3] == 'ИПМ':
        return 'Прикладная математика и информатика'
    if sstr[0:3] == 'ИФИ':
        return 'Фундаментальная информатика и информационные технологии'


def get_cipher(sstr):
    if sstr[0:3] == 'ИУС' and sstr[3] == 'б':
        return '27.03.04'
    if sstr[0:3] == 'ИУС' and sstr[3] == 'м':
        return '27.04.04'
    if sstr[0:3] == 'ИФИ' and sstr[3] == 'м':
        return '02.04.02'
    if sstr[0:3] == 'ИПМ' and sstr[3] == 'б':
        return '01.03.02'
    if sstr[0:3] == 'ИПМ' and sstr[3] == 'м':
        return '01.04.02'


def get_dep_index(f_col):
    try:
        code = df[f_col][mpu_list[j][0]]
    except Exception:
        code = df[f_col][mpu_list[j][0]]
    subject_type = code.strip().split('.')
    if subject_type[1] == 'О':
        return 'Базовая компонента'
    if subject_type[1] == 'В' and len(subject_type) == 3:
        return 'Вариативная компонента'
    if subject_type[1] == 'В' and subject_type[2] == 'ДВ':
        return 'Дисциплины по выбору'
    if subject_type[1] == 'В' and subject_type[2] == 'КР':
        return 'Курсовые работы / проекты'
    if subject_type[0] == 'Б2':
        return 'Практики и НИР'
    if subject_type[0] == 'Б3':
        return 'Государственная итоговая аттестация'
    if subject_type[0] == 'ФТД':
        return 'Факультативы'


mpu_list = get_dep()
counter = 0
course_counter = 0
for i in range(len(col_names)):
    counter += 1
    course_counter += 1
    # print(counter)
    for j in range(len(mpu_list)):
        if df[col_names[i]][mpu_list[j][0]] is not np.nan:
            if counter == 1:
                if course_counter == 5:
                    python_dict['course'] = get_course(
                        str(df[col_names[i-4]][4]))
                else:
                    python_dict['course'] = get_course(
                        str(df[col_names[i]][4]))

                python_dict['name'] = df['Unnamed: 2'][mpu_list[j][0]]
                python_dict['holding_type'] = to_full_holding_type(
                    df[col_names[i]][8])
                python_dict['total_time_for_group'] = df[col_names[i]
                                                         ][mpu_list[j][0]]
                python_dict['department'] = mpu_list[i][1]
                python_dict['semester'] = get_sem(str(df[col_names[i]][6]))
                python_dict['semester_duration'] = get_sem_dur(
                    str(df[col_names[i]][6]))
                python_dict['credit'] = df[col_names[i+3]][mpu_list[j][0]]
                python_dict['subject_type'] = get_dep_index(f_col)
                python_dict['group_name'] = get_group_name(df[col_names[0]][1])
                python_dict['study_level'] = get_study_level(
                    df[col_names[0]][1])
                python_dict['direction'] = get_direction(
                    python_dict['group_name'])
                python_dict['cipher'] = get_cipher(python_dict['group_name'])

                # print(python_dict)
                list_of_dicts.append(copy(python_dict))
                python_dict.clear()

            if counter == 2:
                if course_counter == 6:
                    python_dict['course'] = get_course(
                        str(df[col_names[i-5]][4]))
                else:
                    python_dict['course'] = get_course(
                        str(df[col_names[i-1]][4]))

                python_dict['name'] = df['Unnamed: 2'][mpu_list[j][0]]
                python_dict['holding_type'] = to_full_holding_type(
                    df[col_names[i]][8])
                python_dict['total_time_for_group'] = df[col_names[i]
                                                         ][mpu_list[j][0]]
                python_dict['department'] = mpu_list[i][1]
                python_dict['semester'] = get_sem(
                    str(df[col_names[i-1]][6]))
                python_dict['semester_duration'] = get_sem_dur(
                    str(df[col_names[i-1]][6]))
                python_dict['credit'] = df[col_names[i+2]][mpu_list[j][0]]
                python_dict['subject_type'] = get_dep_index(f_col)
                python_dict['group_name'] = get_group_name(df[col_names[0]][1])
                python_dict['study_level'] = get_study_level(
                    df[col_names[0]][1])
                python_dict['direction'] = get_direction(
                    python_dict['group_name'])
                python_dict['cipher'] = get_cipher(python_dict['group_name'])

                # print(python_dict)
                list_of_dicts.append(copy(python_dict))
                python_dict.clear()

            if counter == 3:
                if course_counter == 7:
                    python_dict['course'] = get_course(
                        str(df[col_names[i-6]][4]))
                else:
                    python_dict['course'] = get_course(
                        str(df[col_names[i-2]][4]))

                python_dict['name'] = df['Unnamed: 2'][mpu_list[j][0]]
                python_dict['holding_type'] = to_full_holding_type(
                    df[col_names[i]][8])
                python_dict['total_time_for_group'] = df[col_names[i]
                                                         ][mpu_list[j][0]]
                python_dict['department'] = mpu_list[i][1]
                python_dict['semester'] = get_sem(str(df[col_names[i-2]][6]))
                python_dict['semester_duration'] = get_sem_dur(
                    str(df[col_names[i-2]][6]))
                python_dict['credit'] = df[col_names[i+1]][mpu_list[j][0]]
                python_dict['subject_type'] = get_dep_index(f_col)
                python_dict['group_name'] = get_group_name(df[col_names[0]][1])
                python_dict['study_level'] = get_study_level(
                    df[col_names[0]][1])
                python_dict['direction'] = get_direction(
                    python_dict['group_name'])
                python_dict['cipher'] = get_cipher(python_dict['group_name'])

                # print(python_dict)
                list_of_dicts.append(copy(python_dict))
                python_dict.clear()

            if counter == 4:
                counter = 0
            if course_counter == 8:
                course_counter = 0
            # print(df[col_names[i]][mpu_list[j][0]])


print(list_of_dicts)

with open('database_1.json', 'w') as f:
    json.dump(list_of_dicts, f, indent=4)
