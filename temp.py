import pandas as pd
import os


def check_data(df,missed_df):
    """
    Функция для проверки недостающих данных
    :param df: данные по группе
    :return: датафрейм с найденными ошибками
    """

    for row in df.itertuples():
        result_check = ''
        # Проверяем паспортные данные
        result_check += check_passport(row[4:11])
        # Создаем словарь для итерации и проверки отсутствующих значений
        dct_other_data = {'ФИО': f'{row[1]} {row[2]} {row[3]}', 'СНИЛС': row[11], 'ИНН': row[12],
                          'Адрес регистрации': row[13], 'Фактический адрес': row[14],
                          'Телефон': row[15], 'email': row[16], 'Сирота': row[17]
            , 'Малоимущий': row[18], 'СОП': row[19], 'Номер аттестата': row[20], 'Ср.балл': row[21],
                          'Название школы': row[22],
                          'Населенный пункт школы': row[23], 'Год окончания': row[24], 'Год приема': row[25]
            , 'Текущий курс': row[26], 'Группа': row[27], 'Отделение': row[28]}

        for key, value in dct_other_data.items():
            if value == 'НЕ ЗАПОЛНЕНО!!!':
                result_check += f'{key} Не заполнен!,'
                continue
    print(result_check)
    # Добавляем полученные данные в датафрейм
    # Создаем промежуточный датафрейм
    temp_missed_df = pd.DataFrame(
        {'Отделение': dct_other_data['Отделение'], 'Группа': dct_other_data['Группа'], 'ФИО': dct_other_data['ФИО'],
         'Статус': 'Данные корректны' if result_check == '' else result_check},index=[0]
        )
    print(temp_missed_df)
    missed_df = missed_df.append(temp_missed_df)


def check_passport(row: tuple):
    """
    Проверка паспортных данных
    :param row:
    :return:Строку с результатами проверки
    """
    result_check_passport = ''
    # Создаем словарь
    dct_passport = {'Серия': row[0], 'Номер': row[1], 'Код подразделения': row[2], 'Выдан': row[3],
                    'Дата выдачи': row[4],
                    'Дата рождения': row[5], 'Место рождения': row[6]}
    # Проверяем пустые значения
    for key, value in dct_passport.items():
        if value == 'НЕ ЗАПОЛНЕНО!!!':
            result_check_passport += f'{key} Не заполнен!,'
            continue
    if len(str(dct_passport['Серия'])) != 4 or len(str(dct_passport['Номер'])) != 6:
        result_check_passport += f'Некорректные Серия или Номер паспорта.'
    return result_check_passport


# Путь к файлам котороые нужно соединить
path = 'resources/MO/'

# Базовый файл куда будут добавлятся данные
base_df = pd.read_excel('resources/base.xlsx')
# Файл с ошибочными данными
missed_df = pd.read_excel('resources/missed_data.xlsx')
# Обработка файлов
# Перебираем файлы в указанной директории. Создаем датафрейм из указанного листа
for file in os.listdir(path):
    temp_df = pd.read_excel(f'{path}/{file}', sheet_name='Список')
    temp_df.fillna('НЕ ЗАПОЛНЕНО!!!', inplace=True)
    # Функция для проверки данных
    check_data(temp_df,missed_df)
    base_df = base_df.append(temp_df)
# Вставляем столбец после ФИО, что логично
base_df.insert(3, 'Наименование документа', 'Паспорт гражданина Российской Федерации')

print(missed_df.head())
missed_df.to_excel('Некорректные данные.xlsx',index=False)
base_df.to_excel('base_to_import.xlsx', index=False)
