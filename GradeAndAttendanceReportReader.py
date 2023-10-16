# -*- coding: utf-8 -*-

import pandas as pd
from pathlib import Path

trimester = 'T3'

core_class_names = {'W-LIT11': 'English', 'BIO11': 'Biology', 'WHISTORY11': 'History', 'IMP3': 'Math'}

elective_class_names = {'E-POLITICS': 'Politics-E',
                        'E-KEYBRD1': 'Keyboarding-E',
                        'E-CHORUS': 'Chorus-E',
                        'E-INDSOCST': 'Ind. Social Studies-E',
                        'E-CENS&SOC': 'Censorship-E',
                        'E-SHRT-FIC': 'Short Stories-E',
                        'E-EL-ART1': 'Art-E',
                        'E-ARTSTDIO': 'Art Studio-E',
                        'E-TRIG': 'Trigonometry-E',
                        'E-PRE-CALC': 'Pre-Calc-E',
                        'E-BOTANY': 'Botany-E',
                        'E-ROBOTICS': 'Robotics-E',
                        'E-DE-ADOBE': 'Adobe-E',
                        'E-DE-MNFCT': 'Manufacturing-E',
                        'E-ACCOUNT': 'Accounting-E',
                        'E-ECONOMIC': 'Economics-E',
                        'INTERNSHIP': 'Internship',
                        'E-FIT11': 'Fitness-E',
                        'E-HEALTH11': 'Health-E',
                        'CAREER': 'Career'}

other_class_names = {'TECH11': 'Technology', 'FI-LIT2': 'Fin. Lit.', 'WRKFORCE11': 'Workforce', 'ESL': 'ESL',
                     'GRAD-PROJ1': 'Grad. Project', 'Average': 'Average',
                     'Fail Count': 'Fail Count'}

class_names = dict(list(core_class_names.items()) + list(elective_class_names.items()) + list(other_class_names.items()))

student_dd = {}
all_weeks_list = []


def add_row(stu_name, cl_name, data_dict):
    return_rows = []
    new_row = [student_dd[stu_name]['LS'], student_dd[stu_name]['Team'], stu_name, student_dd[stu_name]['Advisor'],
               cl_name]
    for w in range(all_weeks_list[-1], all_weeks_list[0] - 1, -1):
        if cl_name in student_dd[stu_name][data_dict].keys():
            if w in student_dd[stu_name][data_dict][cl_name].keys():
                new_row.append(student_dd[stu_name][data_dict][cl_name][w])
            else:
                if data_dict == "Attendance":
                    new_row.append(0)
                else:
                    new_row.append('-')
        else:
            if data_dict == "Attendance":
                new_row.append(0)
            else:
                new_row.append('-')

    return_rows.append(new_row)

    if data_dict == 'Attendance':
        dat_row = new_row[:5]
        if cl_name in student_dd[stu_name]['Grades'].keys():
            if all_weeks_list[-1] in student_dd[stu_name]['Grades'][cl_name].keys():
                dat_row.append(student_dd[stu_name]['Grades'][cl_name][all_weeks_list[-1]])
            else:
                dat_row.append('-')

            if cl_name not in student_dd[stu_name]['Attendance'].keys():
                dat_row.append(0)
            else:
                dat_row.append(new_row[5])
        else:
            print('how did i get here?')

        return_rows.append(dat_row)

    return return_rows


if __name__ == '__main__':

    sorted_class_names = ['Accounting-E',
                          'Adobe-E',
                          'Art-E',
                          'Art Studio-E',
                          'Biology',
                          'Botany-E',
                          'Career',
                          'Censorship-E',
                          'Chorus-E',
                          'Economics-E',
                          'ESL',
                          'English',
                          'Fin. Lit.',
                          'Fitness-E',
                          'Grad. Project',
                          'Health-E',
                          'History',
                          'Ind. Social Studies-E',
                          'Internship',
                          'Keyboarding-E',
                          'Math',
                          'Manufacturing-E',
                          'Politics-E',
                          'Pre-Calc-E',
                          'Robotics-E',
                          'Short Stories-E',
                          'Technology',
                          'Trigonometry-E',
                          'Workforce',
                          'Average',
                          'Fail Count']

    grade_reports_directory = Path.cwd() / 'Weekly Excel Reports' / 'Grades'

    for xlsx_file_path in grade_reports_directory.glob('*.xlsx'):
        week_num = int(xlsx_file_path.stem.split(' ')[1])
        if week_num not in all_weeks_list:
            all_weeks_list.append(week_num)
        grades_df = pd.read_excel(grade_reports_directory / xlsx_file_path, engine='openpyxl', header=1)

        current_name = ''
        current_fail_count = 0
        current_class_count = 0
        current_grade_total = 0
        for index, student_row in grades_df.iterrows():
            if not pd.isnull(student_row['Homeroom']) and student_row['Homeroom'][:2] == '11' and student_row['Marking Period'][:2] == trimester:
                name = student_row['Student Name']
                if len(current_name) == 0:
                    current_name = name
                team = student_row['Homeroom'][2]
                section = student_row['Section'].split(' ')[1]
                grade = float(student_row['Average'])
                advisor = student_row['Staff Advisor'].split(' ')[1]
                if advisor == 'Luft':
                    advisor = 'Hoskey/Luft'
                elif advisor == 'Seidler':
                    advisor = 'Seidler/Honkala'
                elif advisor == 'Boice':
                    advisor = 'Boice/Fernandes'
                if pd.isna(student_row['*Special Ed?']):
                    ls = ''
                else:
                    ls = student_row['*Special Ed?']

                if name not in student_dd.keys():
                    student_dd[name] = {}
                    student_dd[name]['Advisor'] = advisor
                    student_dd[name]['Team'] = team
                    student_dd[name]['LS'] = ls
                    student_dd[name]['Grades'] = {}
                    student_dd[name]['Attendance'] = {}

                if name != current_name and len(current_name) > 0:
                    if 'Average' not in student_dd[current_name]['Grades'].keys():
                        student_dd[current_name]['Grades']['Average'] = {}
                    if 'Fail Count' not in student_dd[current_name]['Grades'].keys():
                        student_dd[current_name]['Grades']['Fail Count'] = {}
                    try:
                        student_dd[current_name]['Grades']['Average'][week_num] = current_grade_total / current_class_count
                    except ZeroDivisionError:
                        print("{} has no classes".format(current_name))
                        student_dd[current_name]['Grades']['Average'][week_num] = 0
                    student_dd[current_name]['Grades']['Fail Count'][week_num] = current_fail_count
                    current_name = name
                    current_fail_count = 0
                    current_class_count = 0
                    current_grade_total = 0

                if class_names[section] not in student_dd[name]['Grades'].keys():
                    student_dd[name]['Grades'][class_names[section]] = {}

                student_dd[name]['Grades'][class_names[section]][week_num] = grade
                if class_names[section] != 'Workforce':
                    current_class_count += 1
                    current_grade_total += grade
                    if grade < 70:
                        current_fail_count += 1

        #get the last student
        if 'Average' not in student_dd[current_name]['Grades'].keys():
            student_dd[current_name]['Grades']['Average'] = {}
        if 'Fail Count' not in student_dd[current_name]['Grades'].keys():
            student_dd[current_name]['Grades']['Fail Count'] = {}
        student_dd[current_name]['Grades']['Average'][week_num] = current_grade_total / current_class_count
        student_dd[current_name]['Grades']['Fail Count'][week_num] = current_fail_count
        current_name = name
        current_fail_count = 0
        current_class_count = 0
        current_grade_total = 0

    att_reports_directory = Path.cwd() / 'Weekly Excel Reports' / 'Attendance'

    for xlsx_file_path in att_reports_directory.glob('*.xlsx'):
        week_num = int(xlsx_file_path.stem.split(' ')[1])
        attendance_df = pd.read_excel(att_reports_directory / xlsx_file_path, engine='openpyxl', header=1)

        for index, student_row in attendance_df.iterrows():
            if student_row['Homeroom'][:2] == '11' and student_row['Marking Period'][:2] == trimester:
                name = student_row['Student Name']
                section = student_row['Section'].split(' ')[1]
                if len(current_name) == 0:
                    current_name = name
                team = student_row['Homeroom'][2]
                try:
                    advisor = student_row['Staff Advisor'].split(' ')[1]
                except KeyError:
                    try:
                        advisor = student_dd[name]['Advisor']
                    except KeyError:
                        advisor = "unknown"
                if advisor == 'Luft':
                    advisor = 'Hoskey/Luft'
                elif advisor == 'Seidler':
                    advisor = 'Seidler/Honkala'
                elif advisor == 'Boice':
                    advisor = 'Boice/Fernandes'
                try:
                    if pd.isna(student_row['*Special Ed?']):
                        ls = ''
                    else:
                        ls = student_row['*Special Ed?']
                except KeyError:
                    ls = student_dd[name]['LS']

                if name not in student_dd.keys():
                    student_dd[name] = {}
                    student_dd[name]['Advisor'] = advisor
                    student_dd[name]['Team'] = team
                    student_dd[name]['LS'] = ls
                    student_dd[name]['Grades'] = {}
                    student_dd[name]['Attendance'] = {}

                addAU = 0
                if student_row['Attendance Code Name'] == 'AU':
                    addAU = 1

                if name not in student_dd.keys():
                    print(name + ' in attendance report but not in grades report.')

                if class_names[section] not in student_dd[name]['Attendance'].keys():
                    student_dd[name]['Attendance'][class_names[section]] = {week_num: 0}
                elif week_num not in student_dd[name]['Attendance'][class_names[section]].keys():
                    student_dd[name]['Attendance'][class_names[section]][week_num] = 0
                else:
                    pass

                student_dd[name]['Attendance'][class_names[section]][week_num] += addAU

                if class_names[section] not in student_dd[name]['Grades'].keys():
                    student_dd[name]['Grades'][class_names[section]] = {}

    all_weeks_list.sort()
    grade_rows = []
    att_rows = []
    dat_rows = []
    non_att_class = ['Workforce', 'Average', 'Fail Count']
    for name in sorted(student_dd.keys()):

        if name != "Simon, Tyrese" and name != "Rivera, David":
            for class_name in sorted_class_names:
                if class_name in student_dd[name]['Grades'].keys():
                    grade_row = add_row(name, class_name, 'Grades')[0]
                    grade_rows.append(grade_row)
                elif class_name in student_dd[name]['Attendance'].keys():
                    grade_row = add_row(name, class_name, 'Grades')[0]
                    grade_rows.append(grade_row)
                else:
                    pass

            class_count = 0
            for class_name in sorted_class_names:
                if class_name in student_dd[name]['Grades'].keys() and class_name not in non_att_class:
                    rows = add_row(name, class_name, 'Attendance')
                    att_rows.append(rows[0])
                    dat_rows.append(rows[1])
                    class_count += 1
                elif class_name in student_dd[name]['Attendance'].keys():
                    rows = add_row(name, class_name, 'Attendance')
                    att_rows.append(rows[0])
                    dat_rows.append(rows[1])

                    class_count += 1
                else:
                    pass

            while class_count < 6:
                dat_rows.append(['-','-','-','-','-','-','-'])
                class_count += 1

    grades_csv_name = input("What do you want to call the .csv file where the grade results will be written to?\n")
    grades_csv_name += '.csv'
    grades_csv_path = Path.cwd() / grades_csv_name
    grades_csv_path.write_text('')
    grades_csv = grades_csv_path.open('a')

    att_csv_name = input("What do you want to call the .csv file where the attendance results will be written to?\n")
    att_csv_name += '.csv'
    att_csv_path = Path.cwd() / att_csv_name
    att_csv_path.write_text('')
    att_csv = att_csv_path.open('a')

    dat_csv_name = input("What do you want to call the .csv file where the DAT results will be written to?\n")
    dat_csv_name += '.csv'
    dat_csv_path = Path.cwd() / dat_csv_name
    dat_csv_path.write_text('')
    dat_csv = dat_csv_path.open('a')

    grades_csv.write('LS$Team$Name$Advisor$Class')
    att_csv.write('LS$Team$Name$Advisor$Class')
    dat_csv.write('LS$Team$Name$Advisor$Class$Week ' + str(all_weeks_list[-1]) + ' Grade$AUs Through Week '
                  + str(all_weeks_list[-1]) + '\n')
    for week in range(all_weeks_list[-1], all_weeks_list[0]-1, -1):
        grades_csv.write('$Week ' + str(week))
        att_csv.write('$Week ' + str(week))
    grades_csv.write('\n')
    att_csv.write('\n')

    for row in grade_rows:
        str_row = []
        for i in row:
            if row[4] == 'Average' and row.index(i) > 4:
                try:
                    str_row.append(str(round(i, 2)))
                except TypeError:
                    str_row.append("-")
            else:
                str_row.append(str(i))
        grades_csv.write('$'.join(str_row))
        if row != grade_rows[-1]:
            grades_csv.write('\n')

    for row in att_rows:
        str_row = []
        for i in row:
            str_row.append(str(i))
        att_csv.write('$'.join(str_row))
        if row != att_rows[-1]:
            att_csv.write('\n')

    for row in dat_rows:
        str_row = []
        for i in row:
            str_row.append(str(i))
        dat_csv.write('$'.join(str_row))
        if row != dat_rows[-1]:
            dat_csv.write('\n')

    grades_csv.close()
    att_csv.close()
    dat_csv.close()
