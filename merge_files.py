import pandas as pd
import os

# Отражение максимального количества колонок в пайчарме
pd.set_option('display.max_columns', None)


def create_table_groups_to_import(df):
    """
    Функция для создания таблицы с группами
    :param df:базовый датафрейм со всеми данными
    :return:Создает таблицу установленного вида и формата
    """
    edu_df = pd.read_excel('resources/edu_ou_t.xlsx')

    # Нам нужны только id специальностей, поэтому создаем словарь вида код специальности:id специальности
    dct_ou = dict()

    for row in edu_df.itertuples():
        dct_ou[row[4]] = row[1]
    # Мы получили словарь с нужными айди
    #
    # Группируем
    group_df = df.groupby('Группа')
    out_df_group = group_df.first()
    # Выносим группы из индекса
    out_df_group.reset_index(inplace=True)
    # Забираем нужные столбцы
    out_df_group = out_df_group[['Код специальности', 'Группа', 'Год приема в БРИТ', 'Текущий курс']]
    # Переименовываем их
    out_df_group.columns = ['short_title_p', 'Группа', 'Год приема в БРИТ', 'Текущий курс']
    out_df_group.loc[:, 'id'] = None

    for row in out_df_group.itertuples():
        out_df_group.loc[row[0], 'id'] = dct_ou[row[1]]
    # Создаем итоговый датафрейм
    group_t_columns = pd.DataFrame(index=range(1, out_df_group.shape[0]))
    group_t_columns['edu_ou_id'] = out_df_group['id']
    group_t_columns['name_p'] = out_df_group['Группа']
    group_t_columns['year_p'] = out_df_group['Год приема в БРИТ']
    group_t_columns['course_p'] = out_df_group['Текущий курс']
    group_t_columns.to_excel('group_t_columns.xlsx')


def create_table_person_to_import(df):
    """
    Функция для генерации таблицы импорта персон
    :param df:
    :return:
    """
    # Считываем базовый датафрейм с колонками
    person_df = pd.read_excel('resources/person_t_columns.xlsx')
    # Количество строк
    count_person = person_df.shape[0]
    # Начинаем заполнять таблицу
    person_df['id'] = range(1, count_person + 1)
    person_df['identity_type_p'] = '1 Паспорт гражданина Российской Федерации'
    for row in df.itertuples():
        person_df.loc[row[0], 'identity_seria'] = row[4]
        person_df.loc[row[0], 'identity_number'] = row[5]
        person_df.loc[row[0], 'identity_firstName'] = row[2]
        person_df.loc[row[0], 'identity_lastName'] = row[1]
        person_df.loc[row[0], 'identity_middleName'] = row[3]
        person_df.loc[row[0], 'identity_birthDate'] = row[9]
        person_df.loc[row[0], 'identity_birthPlace'] = row[10]
        person_df.loc[row[0], 'identity_sex_p'] = row[14]
        person_df.loc[row[0], 'identity_date_p'] = row[8]
        person_df.loc[row[0], 'identity_middleName'] = row[3]
        person_df.loc[row[0], 'identity_code_p'] = row[6]
        person_df.loc[row[0], 'identity_place_p'] = row[7]
        person_df.loc[row[0], 'identity_citizenship_p'] = '0 Россия'
        person_df.loc[row[0], 'identity_middleName'] = row[3]
        person_df.loc[row[0], 'identity_middleName'] = row[3]


def check_data(df):
    """
    Функция для проверки недостающих данных
    :param df: данные по группе
    :return: датафрейм с найденными ошибками
    """

    missed_df = pd.DataFrame({'Отделение': None, 'Группа': None, 'ФИО': None, 'Статус': None}, index=['a'])
    for row in df.itertuples():
        result_check = ''
        # Проверяем паспортные данные
        result_check += check_passport(row[4:11])
        # Создаем словарь для итерации и проверки отсутствующих значений
        dct_other_data = {'ФИО': f'{row[1]} {row[2]} {row[3]}', 'СНИЛС': row[11], 'ИНН': str(row[12]),
                          'Гражданство': row[13], 'Пол': row[14],
                          'Адрес регистрации': row[15], 'Фактический адрес': row[16],
                          'Телефон': str(row[17]), 'email': row[18], 'Статус здоровья': row[19], 'Сирота': row[21],
                          'Малоимущий': row[22], 'СОП': row[23],
                          'Вид аттестата': row[24], 'Номер аттестата': row[25], 'Ср.балл': row[26],
                          'Название школы': row[27],
                          'Населенный пункт школы': row[28], 'Год окончания': row[29], 'Год приема': row[30],
                          'База образования': row[31]
            , 'Текущий курс': row[32], 'Группа': row[33], 'Отделение': row[34]}

        for key, value in dct_other_data.items():
            if value == 'НЕ ЗАПОЛНЕНО!!!':
                result_check += f'{key} Не заполнен!,'
                continue
        # Добавляем полученные данные в датафрейм
        # Создаем промежуточный датафрейм
        temp_missed_df = pd.DataFrame([
            {'Отделение': dct_other_data['Отделение'], 'Группа': dct_other_data['Группа'], 'ФИО': dct_other_data['ФИО'],
             'Статус': 'Данные корректны' if result_check == '' else result_check}]
        )
        missed_df = missed_df.append(temp_missed_df)
    return missed_df


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
    # print(f'{key} {value}')
    if len(str(dct_passport['Серия'])) != 4 or len(str(dct_passport['Номер'])) != 6:
        result_check_passport += f'Некорректные Серия или Номер паспорта.'
    return result_check_passport


# Путь к файлам котороые нужно соединить
path = 'resources/data/'

# Базовый файл куда будут добавлятся данные
base_df = pd.read_excel('resources/base.xlsx')
# Файл с ошибочными данными
missed_df = pd.read_excel('resources/missed_data.xlsx')
# Обработка файлов
# Перебираем файлы в указанной директории. Создаем датафрейм из указанного листа
for file in os.listdir(path):
    current_file = file
    try:

        print(file)
        temp_df = pd.read_excel(f'{path}/{file}', sheet_name=0,
                                dtype={'ИНН': str, 'Телефон': str, 'Номер паспорта': str,
                                       'Серия паспорта': str})
        temp_df.fillna('НЕ ЗАПОЛНЕНО!!!', inplace=True)
        # Функция для проверки данных
        missed_df = missed_df.append(check_data(temp_df))
        base_df = base_df.append(temp_df)
    except  KeyError as e:
        with open('errors.txt', 'a', encoding='utf-8') as f:
            f.write(f'{e} {current_file}\n')
        continue
    except:
        with open('errors.txt', 'a', encoding='utf-8') as f:
            f.write(f'{current_file}\n')
        continue

# Вставляем столбец после ФИО, что логично
# base_df.insert(3, 'Наименование документа', 'Паспорт гражданина Российской Федерации')
# create_table_groups_to_import(base_df)
# create_table_person_to_import(base_df)

# Создание таблицы с группами
print(base_df.head())

missed_df.to_excel('Некорректные данные.xlsx', index=False)
base_df.to_excel('base_to_import.xlsx', index=False)
