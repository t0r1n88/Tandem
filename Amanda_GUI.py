import tkinter
import sys
import pandas as pd
import openpyxl
import numpy as np
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from openpyxl.styles import Alignment
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import time
import datetime
# pd.options.mode.chained_assignment = None  # default='warn'
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
pd.options.mode.chained_assignment = None  # default='warn'



def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller
    Функция чтобы логотип отображался"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def select_end_folder():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder_report
    path_to_end_folder_report = filedialog.askdirectory()


def select_end_folder_machine():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder
    path_to_end_folder = filedialog.askdirectory()


def select_file_data_abitur():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global name_file_abiturs
    # Получаем путь к файлу
    name_file_abiturs = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_file_data_person():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global name_file_person
    # Получаем путь к файлу
    name_file_person = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

"""
Функции для машинистов
"""
def select_file_data_abitur_machine():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global path_to_person
    # Получаем путь к файлу
    path_to_person = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def select_file_data_divde():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global path_to_machine
    # Получаем путь к файлу
    path_to_machine = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def select_file_data_reit():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global path_to_reit
    # Получаем путь к файлу
    path_to_reit = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def processing_reit_machine():
    try:
        machine_df = pd.read_excel(path_to_machine)  # распредление по специальностям
        reit_df = pd.read_excel(path_to_reit, skiprows=4, header=None)  # файл с таблицей из ворда
        df_person = pd.read_excel(path_to_person, sheet_name='Абитуриенты', skiprows=8)  # данные абитуриентов

        # генерируем текущее время
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)

        # создаем два датафрейма для электровозников и тепловозников
        teplo_df = machine_df[['тепловоз', 'атт.1', 'мед справка.1']]
        teplo_df.dropna(axis=0, inplace=True)

        # заполняем пустые строки
        teplo_df['тепловоз'] = teplo_df['тепловоз'].fillna('Не заполнено ФИО')

        # очищаем от пробельных символов
        teplo_df['тепловоз'] = teplo_df['тепловоз'].apply(lambda x: x.strip())
        elect_df = machine_df[['электровоз', 'атт', 'мед справка']]

        elect_df['электровоз'] = elect_df['электровоз'].fillna('Не заполнено ФИО')
        elect_df['электровоз'] = elect_df['электровоз'].apply(lambda x: x.strip())

        # создаем файл с общим списком
        temp_lst = teplo_df['тепловоз'].tolist()
        temp_lst.extend(elect_df['электровоз'].tolist())

        temp_df = pd.DataFrame(columns=['ФИО общее'])

        temp_df['ФИО общее'] = temp_lst

        snils_df = df_person[['ФИО', 'СНИЛС']]

        snils_df = snils_df.drop_duplicates(subset='ФИО')  # убираем дубликаты

        union_df = snils_df.merge(reit_df, how='outer', left_on='СНИЛС', right_on=1,
                                  indicator=True)  # объединяем датафреймы

        # сохраняем тех  у кого нет снилс
        not_snils = union_df[union_df['_merge'] == 'right_only']  # сохраяем тех у кого нет снилса
        not_snils.rename(columns={1: 'Личный номер'}, inplace=True)
        not_snils = not_snils[['Личный номер']]

        clean_reit = union_df[union_df['_merge'] == 'both']  # отбираем тех кто есть в обоих датафреймах

        clean_reit = clean_reit[['ФИО', 'СНИЛС', 2, 5]]

        # Ищем тех кого нет в файле приемки
        missing_priemka = clean_reit.merge(temp_df, how='outer', left_on='ФИО', right_on='ФИО общее', indicator=True)

        # файл где содержатся ФИО тех кто есть в рейтинге но кого нет в файле приемки
        not_in_priemka = missing_priemka[missing_priemka['_merge'] == 'left_only']
        not_in_priemka.drop(columns=['ФИО общее', '_merge'], inplace=True)
        not_in_priemka.rename(columns={2: 'Приоритет', 5: 'Средний балл'}, inplace=True)

        raw_teplo_df = clean_reit.merge(teplo_df, how='outer', left_on='ФИО', right_on='тепловоз', indicator=True)

        zabr_df_teplo = raw_teplo_df[raw_teplo_df['_merge'] == 'right_only']  # те кто забрал документы
        zabr_df_teplo.drop(columns=['ФИО', 'СНИЛС', 2, 5, '_merge'], inplace=True)

        # электровозы
        raw_electo_df = clean_reit.merge(elect_df, how='outer', left_on='ФИО', right_on='электровоз', indicator=True)
        zabr_df_electo = raw_electo_df[raw_electo_df['_merge'] == 'right_only']  # те кто забрал документы
        zabr_df_electo.drop(columns=['ФИО', 'СНИЛС', 2, 5, '_merge'], inplace=True)

        clean_teplo_df = raw_teplo_df[raw_teplo_df['_merge'] == 'both']  # готовим итоговыйй тепловозник
        clean_teplo_df.drop(columns=['тепловоз', '_merge'], inplace=True)

        clean_teplo_df.columns = ['ФИО', 'СНИЛС', 'Приоритет', 'Средний балл', 'Сдан оригинал', 'Мед.справка']

        clean_teplo_df['Приоритет'] = clean_teplo_df['Приоритет'].astype(int)
        clean_teplo_df.sort_values(by='Средний балл', ascending=False, inplace=True)  # сортируем по убыванию
        snils_teplo_df = clean_teplo_df.drop(columns='ФИО')

        # добавляем правильный индекс
        clean_teplo_df.index = range(1, clean_teplo_df.shape[0] + 1)
        snils_teplo_df.index = range(1, snils_teplo_df.shape[0] + 1)
        clean_teplo_df.index.name = '№'
        snils_teplo_df.index.name = '№'

        # готовим итоговый электровозник
        clean_electo_df = raw_electo_df[raw_electo_df['_merge'] == 'both']  # готовим итоговыйй тепловозник
        clean_electo_df.drop(columns=['электровоз', '_merge'], inplace=True)

        clean_electo_df.columns = ['ФИО', 'СНИЛС', 'Приоритет', 'Средний балл', 'Сдан оригинал', 'Мед.справка']

        clean_electo_df['Приоритет'] = clean_electo_df['Приоритет'].astype(int)
        clean_electo_df.sort_values(by='Средний балл', ascending=False, inplace=True)  # сортируем по убыванию
        snils_electo_df = clean_electo_df.drop(columns='ФИО')

        clean_electo_df.index = range(1, clean_electo_df.shape[0] + 1)
        snils_electo_df.index = range(1, snils_electo_df.shape[0] + 1)
        clean_electo_df.index.name = '№'
        snils_electo_df.index.name = '№'

        with pd.ExcelWriter(f'{path_to_end_folder}/Проверка {current_time}.xlsx') as writer:
            not_snils.to_excel(writer, sheet_name='Нет СНИЛС', index=False)
            not_in_priemka.to_excel(writer, sheet_name='Нет в вашем файле', index=False)
            zabr_df_teplo.to_excel(writer, sheet_name='Тепловоз,нет в рейтинге', index=False)
            zabr_df_electo.to_excel(writer, sheet_name='Электровоз,нет в рейтинге', index=False)

        with pd.ExcelWriter(f'{path_to_end_folder}/Рейтинговые списки Электровоз {current_time}.xlsx') as writer:
            snils_electo_df.to_excel(writer, sheet_name='СНИЛС')
            clean_electo_df.to_excel(writer, sheet_name='ФИО')

        with pd.ExcelWriter(f'{path_to_end_folder}/Рейтинговые списки Тепловоз {current_time}.xlsx') as writer:
            snils_teplo_df.to_excel(writer, sheet_name='СНИЛС')
            clean_teplo_df.to_excel(writer, sheet_name='ФИО')

    except NameError:
        messagebox.showerror('ЦОПП Бурятия Создание отчета приемной комиссии ver 1.7',
                             'Выберите файлы для обработки и конечную папку!')
    else:
        messagebox.showinfo('ЦОПП Бурятия Создание отчета приемной комиссии ver 1.7',
                            'Создание отчета успешно завершено!')


def processing_report():
    """
    Фугкция для обработки данных
    :return:
    """
    try:
        # создаем датафрейм со специальностями
        lst_code = ['23.01.09', '43.01.06', '23.02.06', '43.02.06', '15.01.05', '15.01.35', '15.01.33', '23.01.10'
            , '08.01.31', '23.01.17', '08.02.09', '23.02.07', '13.02.07', '35.01.27']

        lst_name_spec = ['Машинист локомотива', 'Проводник на железнодорожном транспорте',
                         'Техническая эксплуатация подвижного состава железных дорог',
                         'Сервис на транспорте (по видам транспорта)',
                         'Сварщик (ручной и частично механизированной сварки (наплавки)',
                         'Мастер слесарных работ', 'Токарь на станках с числовым программным управлением',
                         'Слесарь по обслуживанию и ремонту подвижного состава'
            , 'Электромонтажник электрических сетей и электрооборудования',
                         'Мастер по ремонту и обслуживанию автомобилей',
                         'Монтаж, наладка и эксплуатация электрооборудования промышленных и гражданских зданий',
                         'Техническое обслуживание и ремонт двигателей, систем и агрегатов автомобилей',
                         'Электроснабжение (по отраслям)',
                         'Мастер сельскохозяйственного производства']

        lst_plan = [100, 25, 25, 25, 50, 25, 25, 50
            , 25, 25, 25, 25, 25, 25]
        base_df = pd.DataFrame(columns=['Код', 'Наименование'])
        base_df['Код'] = lst_code
        base_df['Наименование'] = lst_name_spec
        base_df['Направление подготовки'] = base_df['Код'] + ' ' + base_df['Наименование']
        base_df['База'] = '9 кл.'
        base_df['Количество мест'] = lst_plan

        df_abitur = pd.read_excel(name_file_abiturs, skiprows=3, usecols=['Абитуриент', 'Доп. статус', '№ заявления'])
        df_person = pd.read_excel(name_file_person, sheet_name='Абитуриенты', skiprows=8,
                                  usecols=['ФИО', 'Нуждается в общежитии', 'Формирующее подр.',
                                           'Направление подготовки', 'Сдан оригинал', 'Состояние выбран. конкурса',
                                           'СНИЛС'])

        df_person = df_person[~df_person['Направление подготовки'].isnull()]  # убираем тех у кого нет заявлений
        df_abitur = df_abitur[~df_abitur['№ заявления'].isnull()]


        df_dupl = df_person.drop_duplicates(subset=['ФИО'])  # создаем датафрейм без дубликатов

        dupl_cross_df = df_dupl.merge(df_abitur, how='inner', left_on='ФИО', right_on='Абитуриент')

        # Преобразовываем да-нет в 1 или 0 для подсчетов
        dupl_cross_df['Нуждается в общежитии'] = dupl_cross_df['Нуждается в общежитии'].apply(
            lambda x: 0 if x == 'нет' else 1)
        dupl_cross_df['Сдан оригинал'] = dupl_cross_df['Сдан оригинал'].apply(lambda x: 0 if x == 'нет' else 1)
        # заменяем нан на пустые строки чтобы произвести поиск слова сирота;
        dupl_cross_df['Доп. статус'].fillna('', inplace=True)
        dupl_cross_df['Сироты'] = dupl_cross_df['Доп. статус'].apply(lambda x: 1 if 'Сирота;' in x else 0)
        dupl_cross_df['СВО'] = dupl_cross_df['Доп. статус'].apply(
            lambda x: 1 if 'Дети военнослужащих, участвующих в спецоперации' in x else 0)
        dupl_cross_df['Целевой договор'] = dupl_cross_df['Доп. статус'].apply(
            lambda x: 1 if 'Целевой договор' in x else 0)

        dupl_cross_df['for_counting'] = 1

        dupl_cross_df.drop(columns=['Доп. статус'], inplace=True)

        dupl_svod_df = pd.DataFrame.pivot_table(dupl_cross_df,
                                                index=['Формирующее подр.', 'Направление подготовки'],
                                                values=['Сдан оригинал', 'Сироты', 'СВО', 'Целевой договор',
                                                        'Нуждается в общежитии'],
                                                aggfunc='sum')

        dupl_svod_df.columns = ['Нуждается в общежитии чел.', 'Дети СВО', 'Сдано оригиналов', 'Сирот чел.',
                                'Целевой договор']

        # Меняем местами столбцы
        single_out_df = dupl_svod_df.reindex(
            columns=['Сдано оригиналов',
                     'Нуждается в общежитии чел.',
                     'Сирот чел.', 'Дети СВО', 'Целевой договор'])

        # Соединяем оба датафрейма

        cross_df = df_person.merge(df_abitur, how='inner', left_on='ФИО', right_on='Абитуриент')


        # Преобразовываем да-нет в 1 или 0 для подсчетов
        cross_df['Нуждается в общежитии'] = cross_df['Нуждается в общежитии'].apply(lambda x: 0 if x == 'нет' else 1)
        cross_df['Сдан оригинал'] = cross_df['Сдан оригинал'].apply(lambda x: 0 if x == 'нет' else 1)

        # заменяем нан на пустые строки чтобы произвести поиск слова сирота;
        cross_df['Доп. статус'].fillna('', inplace=True)
        cross_df['Сироты'] = cross_df['Доп. статус'].apply(lambda x: 1 if 'Сирота;' in x else 0)
        cross_df['СВО'] = cross_df['Доп. статус'].apply(
            lambda x: 1 if 'Дети военнослужащих, участвующих в спецоперации' in x else 0)

        cross_df['for_counting'] = 1

        cross_df.drop(columns=['Доп. статус'], inplace=True)

        # Создаем сокращенный датафрейм чтобы добавить его в базовый
        small_df = cross_df[['Направление подготовки', 'Состояние выбран. конкурса', 'for_counting']]

        # объединяем датафреймы
        union_df = base_df.merge(small_df, how='outer', left_on='Направление подготовки',
                                 right_on='Направление подготовки')
        union_df.fillna(0, inplace=True)

        # забранные заявления
        return_z = union_df[union_df['Состояние выбран. конкурса'] == 'Забрал документы']

        base_df_groupby = union_df.groupby(['Направление подготовки']).agg({'for_counting': sum})
        base_df_groupby['for_counting'] = base_df_groupby['for_counting'].apply(int)
        base_df_groupby = base_df_groupby.reset_index()
        base_df_groupby.rename(columns={'for_counting': 'Подано заявлений'}, inplace=True)

        base_df = base_df.merge(base_df_groupby, how='inner', left_on='Направление подготовки',
                                right_on='Направление подготовки')
        base_df.sort_values(by='Подано заявлений', ascending=False, inplace=True)


        base_df.rename(columns={'Наименование': 'Наименование образовательной программы'})
        base_df.drop(columns='Направление подготовки', inplace=True)

        # считаем количество тех кто забрал документы
        cross_df['Забрали заявления'] = cross_df['Состояние выбран. конкурса'].apply(
            lambda x: 1 if x == 'Забрал документы' else 0)
        cross_df['Заявления'] = cross_df['for_counting']

        svod_df = pd.DataFrame.pivot_table(cross_df,
                                           index=['Формирующее подр.', 'Направление подготовки'],
                                           values=['Заявления', 'Забрали заявления', ],
                                           aggfunc='sum')

        svod_df = svod_df.reindex(columns=['Заявления', 'Забрали заявления'])

        svod_df['Итого заявлений'] = svod_df['Заявления'] - svod_df['Забрали заявления']

        out_df = svod_df.reset_index()

        single_out_df = single_out_df.reset_index()

        finish_df = pd.merge(out_df, single_out_df, how='outer')  # объединяем

        wb = openpyxl.Workbook()
        # Переименовываем лист
        sheet = wb['Sheet']
        sheet.title = 'Отчет'

        sum_row = finish_df.sum(axis=0).to_frame().transpose()

        sum_row['Формирующее подр.'] = 'Всего'
        sum_row['Направление подготовки'] = ''

        # объединяем датафреймы

        all_finish_df = pd.concat([finish_df, sum_row], axis=0)

        for r in dataframe_to_rows(all_finish_df, index=False, header=True):
            if len(r) != 1:
                wb['Отчет'].append(r)

        # # Настраиваем выходной файл
        wb['Отчет'].column_dimensions['A'].width = 30
        wb['Отчет'].column_dimensions['B'].width = 50
        wb['Отчет']['B2'].alignment = Alignment(wrap_text=True)
        wb['Отчет'].column_dimensions['C'].width = 20
        wb['Отчет'].column_dimensions['D'].width = 20
        wb['Отчет'].column_dimensions['F'].width = 20
        wb['Отчет'].column_dimensions['G'].width = 20
        wb['Отчет'].column_dimensions['H'].width = 30
        wb['Отчет']['H1'].alignment = Alignment(wrap_text=True)

        # Получаем текущее время для того чтобы использовать в названии
        t = time.localtime()
        current_time = time.strftime('%H_%M_%d_%m', t)
        # Сохраняем итоговый файл
        base_df.to_excel(f'{path_to_end_folder_report}/Количество поданых заявлений {current_time}.xlsx', index=False)
        wb.save(f'{path_to_end_folder_report}/Ежедневный отчет приемной комиссии ГБПОУ БРИТ {current_time}.xlsx')

        # ищем полных тезок
        temp_dupl_df = df_person.drop_duplicates(subset=['ФИО', 'СНИЛС'])

        tezki_df = temp_dupl_df[temp_dupl_df.duplicated(subset='ФИО', keep=False)]

        tezki_df.to_excel(f'{path_to_end_folder_report}/Полные тезки {current_time}.xlsx', index=False)



    except NameError:
        messagebox.showerror('ЦОПП Бурятия Создание отчета приемной комиссии ver 1.7','Выберите файлы для обработки и конечную папку!')
    else:
        messagebox.showinfo('ЦОПП Бурятия Создание отчета приемной комиссии ver 1.7','Создание отчета успешно завершено!')

"""
Функции для проверки наличия людей
"""
def convert_columns_to_str(df, number_columns):
    """
    Функция для конвертации указанных столбцов в строковый тип и очистки от пробельных символов в начале и конце
    """

    for column in number_columns:  # Перебираем список нужных колонок
        try:
            df.iloc[:, column] = df.iloc[:, column].astype(str)
            # Очищаем колонку от пробельных символов с начала и конца
            df.iloc[:, column] = df.iloc[:, column].apply(lambda x: x.strip())
        except IndexError:
            messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.35',
                                 'Проверьте порядковые номера колонок которые вы хотите обработать.')


def convert_params_columns_to_int(lst):
    """
    Функция для конвератации значений колонок которые нужно обработать.
    Очищает от пустых строк, чтобы в итоге остался список из чисел в формате int
    """
    out_lst = [] # Создаем список в который будем добавлять только числа
    for value in lst: # Перебираем список
        try:
            # Обрабатываем случай с нулем, для того чтобы после приведения к питоновскому отсчету от нуля не получилась колонка с номером -1
            number = int(value)
            if number != 0:
                out_lst.append(value) # Если конвертирования прошло без ошибок то добавляем
            else:
                continue
        except: # Иначе пропускаем
            continue
    return out_lst

def create_doc_convert_date(cell):
    """
    Функция для конвертации даты при создании документов
    :param cell:
    :return:
    """
    try:
        string_date = datetime.datetime.strftime(cell, '%d.%m.%Y')
        return string_date
    except ValueError:
        return 'Не удалось конвертировать дату.Проверьте значение ячейки!!!'
    except TypeError:
        return 'Не удалось конвертировать дату.Проверьте значение ячейки!!!'

def check_date_columns(i, value):
    """
    Функция для проверки типа колонки. Необходимо найти колонки с датой
    :param i:
    :param value:
    :return:
    """
    try:
        itog = pd.to_datetime(str(value), infer_datetime_format=True)
    except:
        pass
    else:
        return i

def processing_date_column(df, lst_columns):
    """
    Функция для обработки столбцов с датами. конвертация в строку формата ДД.ММ.ГГГГ
    """
    # получаем первую строку
    first_row = df.iloc[0, lst_columns]

    lst_first_row = list(first_row)  # Превращаем строку в список
    lst_date_columns = []  # Создаем список куда будем сохранять колонки в которых находятся даты
    tupl_row = list(zip(lst_columns,
                        lst_first_row))  # Создаем список кортежей формата (номер колонки,значение строки в этой колонке)

    for idx, value in tupl_row:  # Перебираем кортеж
        result = check_date_columns(idx, value)  # проверяем является ли значение датой
        if result:  # если да то добавляем список порядковый номер колонки
            lst_date_columns.append(result)
        else:  # иначе проверяем следующее значение
            continue
    for i in lst_date_columns:  # Перебираем список с колонками дат, превращаем их в даты и конвертируем в нужный строковый формат
        df.iloc[:, i] = pd.to_datetime(df.iloc[:, i], errors='coerce', dayfirst=True)
        df.iloc[:, i] = df.iloc[:, i].apply(create_doc_convert_date)

def clean_ending_columns(lst_columns:list,name_first_df,name_second_df):
    """
    Функция для очистки колонок таблицы с совпадающими данными от окончаний _x _y

    :param lst_columns:
    :param time_generate
    :param name_first_df
    :param name_second_df
    :return:
    """
    out_columns = [] # список для очищенных названий
    for name_column in lst_columns:
        if '_x' in name_column:
            # если они есть то проводим очистку и добавление времени
            cut_name_column = name_column[:-2] # обрезаем
            temp_name = f'{cut_name_column}_{name_first_df}' # соединяем
            out_columns.append(temp_name) # добавляем
        elif '_y' in name_column:
            cut_name_column = name_column[:-2]  # обрезаем
            temp_name = f'{cut_name_column}_{name_second_df}'  # соединяем
            out_columns.append(temp_name)  # добавляем
        else:
            out_columns.append(name_column)
    return out_columns

def select_file_params_comparsion():
    """
    Функция для выбора файла с параметрами колонок т.е. кокие колонки нужно обрабатывать
    :return:
    """
    global file_params
    file_params = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_first_comparison():
    """
    Функция для выбора  первого файла с данными которые нужно сравнить
    :return: Путь к файлу с данными
    """
    global name_first_file_comparison
    # Получаем путь к файлу
    name_first_file_comparison = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_second_comparison():
    """
    Функция для выбора  второго файла с данными которые нужно сравнить
    :return: Путь к файлу с данными
    """
    global name_second_file_comparison
    # Получаем путь к файлу
    name_second_file_comparison = filedialog.askopenfilename(
        filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def select_end_folder_comparison():
    """
    Функция для выбора папки куда будет генерироваться итоговый файл
    :return:
    """
    global path_to_end_folder_comparison
    path_to_end_folder_comparison = filedialog.askdirectory()


def processing_comparison():
    """
    Функция для сравнения 2 колонок
    :return:
    """
    try:
        # Получаем значения текстовых полей
        first_sheet_name = 'Абитуриенты'
        second_sheet_name = 'Отчет'
        # загружаем файлы
        first_df = pd.read_excel(name_first_file_comparison, sheet_name=first_sheet_name, dtype=str,
                                 keep_default_na=False,skiprows=3)
        # получаем имя файла
        name_first_df = name_first_file_comparison.split('/')[-1]
        name_first_df = name_first_df.split('.xlsx')[0]

        second_df = pd.read_excel(name_second_file_comparison, sheet_name=second_sheet_name, dtype=str,
                                  keep_default_na=False,skiprows=3)
        # получаем имя файла
        name_second_df = name_second_file_comparison.split('/')[-1]
        name_second_df = name_second_df.split('.xlsx')[0]

        params = pd.read_excel(file_params, header=None, keep_default_na=False)

        # Преврашаем каждую колонку в список
        params_first_columns = params[0].tolist()
        params_second_columns = params[1].tolist()

        # Конвертируем в инт заодно проверяя корректность введенных данных
        int_params_first_columns = convert_params_columns_to_int(params_first_columns)
        int_params_second_columns = convert_params_columns_to_int(params_second_columns)

        # Отнимаем 1 от каждого значения чтобы привести к питоновским индексам
        int_params_first_columns = list(map(lambda x: x - 1, int_params_first_columns))
        int_params_second_columns = list(map(lambda x: x - 1, int_params_second_columns))

        # Конвертируем нужные нам колонки в str
        convert_columns_to_str(first_df, int_params_first_columns)
        convert_columns_to_str(second_df, int_params_second_columns)



        # Проверяем наличие колонок с датами в списке колонок для объединения чтобы привести их в нормальный вид
        for number_column_params in int_params_first_columns:
            if 'дата' in first_df.columns[number_column_params].lower():
                first_df.iloc[:, number_column_params] = pd.to_datetime(first_df.iloc[:, number_column_params],
                                                                        errors='coerce', dayfirst=True)
                first_df.iloc[:, number_column_params] = first_df.iloc[:, number_column_params].apply(
                    create_doc_convert_date)

        for number_column_params in int_params_second_columns:
            if 'дата' in second_df.columns[number_column_params].lower():
                second_df.iloc[:, number_column_params] = pd.to_datetime(second_df.iloc[:, number_column_params],
                                                                         errors='coerce', dayfirst=True)
                second_df.iloc[:, number_column_params] = second_df.iloc[:, number_column_params].apply(
                    create_doc_convert_date)

        # в этом месте конвертируем даты в формат ДД.ММ.ГГГГ
        # processing_date_column(first_df, int_params_first_columns)
        # processing_date_column(second_df, int_params_second_columns)

        # Проверяем наличие колонки _merge
        if '_merge' in first_df.columns:
            first_df.drop(columns=['_merge'], inplace=True)
        if '_merge' in second_df.columns:
            second_df.drop(columns=['_merge'], inplace=True)
        # Проверяем наличие колонки ID
        if 'ID_объединения' in first_df.columns:
            first_df.drop(columns=['ID_объединения'], inplace=True)
        if 'ID_объединения' in second_df.columns:
            second_df.drop(columns=['ID_объединения'], inplace=True)

        # создаем датафреймы из колонок выбранных для объединения, такой способо связан с тем, что
        # при использовании sum числа в строковом виде превращаются в числа
        key_first_df = first_df.iloc[:,int_params_first_columns]
        key_second_df = second_df.iloc[:,int_params_second_columns]
        # Создаем в каждом датафрейме колонку с айди путем склеивания всех нужных колонок в одну строку
        first_df['ID_объединения'] = key_first_df.apply(lambda x:''.join(x),axis=1)
        second_df['ID_объединения'] = key_second_df.apply(lambda x: ''.join(x), axis=1)


        first_df['ID_объединения'] = first_df['ID_объединения'].apply(lambda x: x.replace(' ', ''))
        second_df['ID_объединения'] = second_df['ID_объединения'].apply(lambda x: x.replace(' ', ''))

        # делаем прописными айди значения по которым будет вестись объединение
        first_df['ID_объединения'] = first_df['ID_объединения'].apply(lambda x: x.upper())
        second_df['ID_объединения'] = second_df['ID_объединения'].apply(lambda x: x.upper())

        # В результат объединения попадают совпадающие по ключу записи обеих таблиц и все строки из этих двух таблиц, для которых пар не нашлось. Порядок таблиц в запросе не

        # Создаем документ
        wb = openpyxl.Workbook()
        # создаем листы
        ren_sheet = wb['Sheet']
        ren_sheet.title = 'Только в тандеме'
        wb.create_sheet(title='Только в госуслугах', index=1)
        wb.create_sheet(title='В обеих таблицах', index=2)
        # wb.create_sheet(title='Обновленная таблица', index=3)
        # wb.create_sheet(title='Объединённая таблица', index=4)


        # Создаем переменные содержащие в себе количество колонок в базовых датареймах
        first_df_quantity_cols = len(first_df.columns)  # не забываем что там добавилась колонка ID

        # Проводим слияние
        itog_df = pd.merge(first_df, second_df, how='outer', left_on=['ID_объединения'], right_on=['ID_объединения'],
                           indicator=True)

        # копируем в отдельный датафрейм для создания таблицы с обновлениями
        update_df = itog_df.copy()

        # Записываем каждый датафрейм в соответсвующий лист
        # Левая таблица
        left_df = itog_df[itog_df['_merge'] == 'left_only']
        left_df.drop(['_merge'], axis=1, inplace=True)

        # Удаляем колонки второй таблицы чтобы не мешались
        left_df.drop(left_df.iloc[:, first_df_quantity_cols:], axis=1, inplace=True)

        # Переименовываем колонки у которых были совпадение во второй таблице, в таких колонках есть добавление _x
        clean_left_columns = list(map(lambda x: x[:-2] if '_x' in x else x, list(left_df.columns)))
        left_df.columns = clean_left_columns
        for r in dataframe_to_rows(left_df, index=False, header=True):
            wb['Только в тандеме'].append(r)

        right_df = itog_df[itog_df['_merge'] == 'right_only']
        right_df.drop(['_merge'], axis=1, inplace=True)

        # Удаляем колонки первой таблицы таблицы чтобы не мешались
        right_df.drop(right_df.iloc[:, :first_df_quantity_cols - 1], axis=1, inplace=True)

        # Переименовываем колонки у которых были совпадение во второй таблице, в таких колонках есть добавление _x
        clean_right_columns = list(map(lambda x: x[:-2] if '_y' in x else x, list(right_df.columns)))
        right_df.columns = clean_right_columns

        for r in dataframe_to_rows(right_df, index=False, header=True):
            wb['Только в госуслугах'].append(r)

        both_df = itog_df[itog_df['_merge'] == 'both']
        both_df.drop(['_merge'], axis=1, inplace=True)
        # Очищаем от _x  и _y
        clean_both_columns = clean_ending_columns(list(both_df.columns), name_first_df, name_second_df)
        both_df.columns = clean_both_columns

        for r in dataframe_to_rows(both_df, index=False, header=True):
            wb['В обеих таблицах'].append(r)

        # Сохраняем общую таблицу
        # Заменяем названия индикаторов на более понятные
        itog_df['_merge'] = itog_df['_merge'].apply(lambda x: 'Данные из первой таблицы' if x == 'left_only' else
        ('Данные из второй таблицы' if x == 'right_only' else 'Совпадающие данные'))
        itog_df['_merge'] = itog_df['_merge'].astype(str)

        clean_itog_df = clean_ending_columns(list(itog_df.columns), name_first_df, name_second_df)
        itog_df.columns = clean_itog_df
        # for r in dataframe_to_rows(itog_df, index=False, header=True):
        #     wb['Объединённая таблица'].append(r)

        # получаем список с совпадающими колонками первой таблицы
        first_df_columns = [column for column in list(update_df.columns) if str(column).endswith('_x')]
        # получаем список с совпадающими колонками второй таблицы
        second_df_columns = [column for column in list(update_df.columns) if str(column).endswith('_y')]
        # Создаем из списка совпадающих колонок второй таблицы словарь, чтобы было легче обрабатывать
        # да конечно можно было сделать в одном выражении но как я буду читать это через 2 недели?
        dct_second_columns = {column.split('_y')[0]: column for column in second_df_columns}

        for column in first_df_columns:
            # очищаем от _x
            name_column = column.split('_x')[0]
            # Обновляем значение в случае если в колонке _merge стоит both, иначе оставляем старое значение,
            # Чтобы обновить значение в ячейке, во второй таблице не должно быть пустого значения или пробела в аналогичной колонке

            update_df[column] = np.where(
                (update_df['_merge'] == 'both') & (update_df[dct_second_columns[name_column]]) & (
                            update_df[dct_second_columns[name_column]] != ' '),
                update_df[dct_second_columns[name_column]], update_df[column])

            # Удаляем колонки с _y
        update_df.drop(columns=[column for column in update_df.columns if column.endswith('_y')], inplace=True)

        # Переименовываем колонки с _x
        update_df.columns = list(map(lambda x: x[:-2] if x.endswith('_x') else x, update_df.columns))

        # удаляем строки с _merge == right_only
        update_df = update_df[update_df['_merge'] != 'right_only']

        # Удаляем служебные колонки
        update_df.drop(columns=['ID_объединения', '_merge'], inplace=True)

        # используем уже созданный датафрейм right_df Удаляем лишнюю колонку в right_df
        right_df.drop(columns=['ID_объединения'], inplace=True)

        # Добавляем нехватающие колонки
        new_right_df = right_df.reindex(columns=update_df.columns, fill_value=None)

        update_df = pd.concat([update_df, new_right_df])

        # for r in dataframe_to_rows(update_df, index=False, header=True):
        #     wb['Обновленная таблица'].append(r)

        # генерируем текущее время
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        # Сохраняем итоговый файл
        wb.save(f'{path_to_end_folder_comparison}/Сравнение 2 таблиц от {current_time}.xlsx')
        # # Сохраняем отдельно обновленную таблицу
        # update_df.to_excel(
        #     f'{path_to_end_folder_comparison}/Таблица с обновленными данными и колонками от {current_time}.xlsx',
        #     index=False)

    except NameError:
        messagebox.showerror('ЦОПП Бурятия Создание отчета приемной комиссии ver 1.7',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
    except KeyError:
        messagebox.showerror('ЦОПП Бурятия Создание отчета приемной комиссии ver 1.7',
                             f'В таблице нет такой колонки!\nПроверьте написание названия колонки')
    except ValueError:
        messagebox.showerror('ЦОПП Бурятия Создание отчета приемной комиссии ver 1.7',
                             f'В таблице нет листа с таким названием!\nПроверьте написание названия листа')

    except FileNotFoundError:
        messagebox.showerror('ЦОПП Бурятия Создание отчета приемной комиссии ver 1.7',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    except:
        messagebox.showerror('ЦОПП Бурятия Создание отчета приемной комиссии ver 1.7',
                             'Возникла ошибка!!! Подробности ошибки в файле error.log')
    else:
        messagebox.showinfo('ЦОПП Бурятия Создание отчета приемной комиссии ver 1.7', 'Данные успешно обработаны')




if __name__ == '__main__':
    window = Tk()
    window.title('ЦОПП Бурятия Создание отчета приемной комиссии ver 1.7')
    window.geometry('700x660')
    window.resizable(False, False)


    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку обработки данных для Приложения 6
    tab_report = ttk.Frame(tab_control)
    tab_control.add(tab_report, text='Ежедневные отчеты')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание образовательных программ
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_report,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nПрограмма для создания отчета приемной комиссии директору\nГБПОУ БРИТ')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img = resource_path('logo.png')

    img = PhotoImage(file=path_to_img)
    Label(tab_report,
          image=img
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать файл с данными абитуриентов
    btn_choose_data_abitur = Button(tab_report, text='1) Выберите файл c\n главной страницы', font=('Arial Bold', 20),
                                    command=select_file_data_abitur
                                    )
    btn_choose_data_abitur.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными персон
    btn_choose_data_person = Button(tab_report, text='2) Выберите файл выборки', font=('Arial Bold', 20),
                                    command=select_file_data_person
                                    )
    btn_choose_data_person.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder = Button(tab_report, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                   command=select_end_folder
                                   )
    btn_choose_end_folder.grid(column=0, row=4, padx=10, pady=10)

    #Создаем кнопку обработки данных

    btn_proccessing_data = Button(tab_report, text='4) Создать отчет', font=('Arial Bold', 20),
                                  command=processing_report
                                  )
    btn_proccessing_data.grid(column=0, row=5, padx=10, pady=10)

    """
    Создаем вкладку для машинистов
    """
    #Создаем вкладку обработки данных для Приложения 6
    tab_machine = ttk.Frame(tab_control)
    tab_control.add(tab_machine, text='Машинисты')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание образовательных программ
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_machine,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nПрограмма для создания отчета приемной комиссии директору\nГБПОУ БРИТ')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img_machine = resource_path('logo.png')

    img_machine = PhotoImage(file=path_to_img_machine)
    Label(tab_machine,
          image=img_machine
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать файл с данными абитуриентов
    btn_choose_data_abitur_machine = Button(tab_machine, text='1) Выберите файл выборки', font=('Arial Bold', 20),
                                    command=select_file_data_abitur_machine
                                    )
    btn_choose_data_abitur_machine.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными разделения на машинистов и электровозников
    btn_choose_data_divide = Button(tab_machine, text='2) Выберите ваш файл', font=('Arial Bold', 20),
                                    command=select_file_data_divde
                                    )
    btn_choose_data_divide.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными рейтинга машинистов
    btn_choose_data_reit = Button(tab_machine, text='3) Файл из рейтинга', font=('Arial Bold', 20),
                                    command=select_file_data_reit
                                    )
    btn_choose_data_reit.grid(column=0, row=4, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder = Button(tab_machine, text='5) Выберите конечную папку', font=('Arial Bold', 20),
                                   command=select_end_folder_machine
                                   )
    btn_choose_end_folder.grid(column=0, row=5, padx=10, pady=10)

    # Создаем кнопку обработки данных

    btn_proccessing_machine = Button(tab_machine, text='6) Создать списки', font=('Arial Bold', 20),
                                  command=processing_reit_machine
                                  )
    btn_proccessing_machine.grid(column=0, row=6, padx=10, pady=10)

    """
    Слияние 2 таблиц
    """
    tab_comparison = ttk.Frame(tab_control)
    tab_control.add(tab_comparison, text='Сравнение с госуслугами')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку Создание документов
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_comparison,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                           '\nДля корректной работы программмы уберите из таблицы объединенные ячейки')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_com = resource_path('logo.png')
    img_comparison = PhotoImage(file=path_com)
    Label(tab_comparison,
          image=img
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_comparison = LabelFrame(tab_comparison, text='Подготовка')
    frame_data_for_comparison.grid(column=0, row=2, padx=10)

    # Создаем кнопку выбрать файл с параметрами
    btn_columns_params = Button(frame_data_for_comparison, text='1) Выберите файл с параметрами сравнения',
                                font=('Arial Bold', 10),
                                command=select_file_params_comparsion)
    btn_columns_params.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку Выбрать  первый файл с данными
    btn_data_first_comparison = Button(frame_data_for_comparison, text='2) Выберите файл с главной страницы',
                                       font=('Arial Bold', 10),
                                       command=select_first_comparison
                                       )
    btn_data_first_comparison.grid(column=0, row=4, padx=10, pady=10)

    # Создаем кнопку Выбрать  второй файл с данными
    btn_data_second_comparison = Button(frame_data_for_comparison, text='4) Выберите файл с госуслуг',
                                        font=('Arial Bold', 10),
                                        command=select_second_comparison
                                        )
    btn_data_second_comparison.grid(column=0, row=7, padx=10, pady=10)


    # Создаем кнопку выбора папки куда будет генерироваьться файл
    btn_select_end_comparison = Button(frame_data_for_comparison, text='5) Выберите конечную папку',
                                       font=('Arial Bold', 10),
                                       command=select_end_folder_comparison
                                       )
    btn_select_end_comparison.grid(column=0, row=10, padx=10, pady=10)

    # Создаем кнопку Обработать данные
    btn_data_do_comparison = Button(tab_comparison, text='6) Обработать данные', font=('Arial Bold', 20),
                                    command=processing_comparison
                                    )
    btn_data_do_comparison.grid(column=0, row=11, padx=10, pady=10)



    window.mainloop()