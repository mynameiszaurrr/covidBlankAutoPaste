import pandas as pd
from docxtpl import DocxTemplate
import datetime

doctor = ""  # Write here doctor name
df = pd.read_excel('')  # Write here excel file way
today = datetime.datetime.today().strftime("%d.%m.%Y")


def nan_checker(element):
    checker = str(element)
    if checker == 'nan':
        return '_'*20
    elif checker == 'NaT':
        return '_' * 20
    else:
        return element


def patient_info(excel_file):
    column_number = 0
    patients_list = []
    for doctor_name in excel_file['Бригада']:
        telephon_number = nan_checker(excel_file.loc[column_number]['Телефон'])
        agree_date = nan_checker(excel_file.loc[column_number]['Дата выдачи согласия'])
        address = nan_checker(excel_file.loc[column_number]['Адрес'])

        if (doctor_name == doctor and 'КГУ' in df.loc[column_number]['Источник']) or (doctor_name == doctor and 'На руках' in df.loc[column_number]['Источник']):
            patients_list.append({
                'Дата выдачи согласия': agree_date,
                'ФИО': f"{excel_file.loc[column_number]['Фамилия']} {excel_file.loc[column_number]['Имя']} {excel_file.loc[column_number]['Отчество']}",
                'Дата рождения': str(excel_file.loc[column_number]['Дата рождения'])[0:10],
                'Адрес': address,
                'Телефон': telephon_number
            })
        column_number += 1
    return patients_list


def number_filter(numbers_list):
    mobile_number = []
    if numbers_list is not float:
        for number in numbers_list.split(' '):
            if number == '':
                continue
            if number[-1].isnumeric():
                pass
            else:
                number = number[:-1]
            if number not in mobile_number and number != '70000000000':
                mobile_number.append(number)
        return ', '.join(mobile_number)
    else:
        return '_____________________'


try:
    ankets_count = 0
    print(f"Всего анкет: {len(df['Бригада'])}")
    for i in patient_info(df):
        doc = DocxTemplate('СогласиеШаблон.docx')
        date_of_birth = f"{i['Дата рождения'][-2:]}.{i['Дата рождения'][-5:-3]}.{i['Дата рождения'][0:4]}"
        if i['Дата выдачи согласия'] == '_' * 20:
            date_of_consent = '_____________________'
        else:
            str_date = str(i['Дата выдачи согласия']).split('-')
            date_of_consent = f"{str_date[-1][:2]}.{str_date[1]}.{str_date[0]}"
        numbers = number_filter(i['Телефон'])
        date_of_birth = str(date_of_birth)
        if date_of_birth == 'aT..NaT':
            date_of_birth = f'___.___.________'
        content = {
            'FIO': i['ФИО'],
            'DataOfBirth': date_of_birth,
            'Address': i['Адрес'],
            'MobileNumber': numbers,
            'DoctorName': doctor,
            'DataOfQuarantie': date_of_consent,
            'TodaData': today
        }
        doc.render(content)
        doc.save(f"finel_patients_documents/{i['ФИО']}.docx")
        ankets_count += 1
        if ankets_count > len(df['Бригада']):
            break
    if ankets_count > 0:
        print("Анкеты созданы! Удачной рабочей смены!")
        print(f'Всего анкет {ankets_count}')
    else:
        print('Доктора нет в анкетах! Проверьте вводные данные!')
except Exception as e:
    print(f"Ошибка {e}!")
