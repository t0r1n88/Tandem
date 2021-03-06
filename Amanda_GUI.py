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
# pd.options.mode.chained_assignment = None  # default='warn'
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


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



def processing_report():
    """
    Фугкция для обработки данных
    :return:
    """
    try:
        df_abitur = pd.read_excel(name_file_abiturs, skiprows=3, usecols=['Абитуриент', 'Доп. статус', 'Состояние'])
        df_person = pd.read_excel(name_file_person, sheet_name='Абитуриенты', skiprows=8,
                                  usecols=['ФИО', 'Нуждается в общежитии', 'Формирующее подр.',
                                           'Направление, специальность, профессия', 'Сдан оригинал'])

        wb = openpyxl.Workbook()
        # Переименовываем лист
        sheet = wb['Sheet']
        sheet.title = 'Отчет'

        # Соединяем оба датафрейма

        cross_df = df_person.merge(df_abitur, how='inner', left_on='ФИО', right_on='Абитуриент')

        # Преобразовываем да-нет в 1 или 0 для подсчетов
        cross_df['Нуждается в общежитии'] = cross_df['Нуждается в общежитии'].apply(lambda x: 0 if x == 'нет' else 1)
        cross_df['Сдан оригинал'] = cross_df['Сдан оригинал'].apply(lambda x: 0 if x == 'нет' else 1)
        cross_df['Состояние'] = cross_df['Состояние'].apply(lambda x: 1 if x == 'Забрал документы' else 0)

        # заменяем нан на пустые строки чтобы произвести поиск слова сирота;
        cross_df['Доп. статус'].fillna('', inplace=True)
        cross_df['Доп. статус'] = cross_df['Доп. статус'].apply(lambda x: 1 if 'Сирота;' in x else 0)

        cross_df['for_counting'] = 1

        svod_df = pd.DataFrame.pivot_table(cross_df,
                                           index=['Формирующее подр.', 'Направление, специальность, профессия'],
                                           values=['for_counting', 'Состояние', 'Сдан оригинал', 'Доп. статус',
                                                   'Нуждается в общежитии'],
                                           aggfunc='sum')

        svod_df.columns = ['Сдали всего', 'Сирот чел.', 'Нуждается в общежитии чел.', 'Сдано оригиналов',
                           'Забрали заявления']

        svod_df['Итого'] = svod_df['Сдали всего'] - svod_df['Забрали заявления']

        # Меняем местами столбцы
        out_df = svod_df.reindex(columns=['Сдали всего', 'Забрали заявления', 'Итого', 'Сдано оригиналов', 'Сирот чел.',
                                          'Нуждается в общежитии чел.'])

        # разворачиваем столбец в строку
        sum_row = out_df.sum(axis=0).to_frame().transpose()

        # Добавляем колонки чтобы сделать из них мультинидекс .Ужасно решение но что есть то есть
        sum_row['1'] = 'Всего'
        sum_row['2'] = ''

        # Делем мультинидекс и объединяем датафреймы
        sum_row.set_index(['1', '2'], inplace=True)
        all_out_df = pd.concat([out_df, sum_row], axis=0)

        # Преобразовываем мультинидекс в колонки
        finish_df = all_out_df.reset_index()

        for r in dataframe_to_rows(finish_df, index=False, header=True):
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
        current_time = time.strftime('%H_%M_%S', t)
        # Сохраняем итоговый файл
        wb.save(f'{path_to_end_folder_report}/Ежедневный отчет приемной комиссии ГБПОУ БРИТ {current_time}.xlsx')
    except NameError:
        messagebox.showerror('ЦОПП Бурятия','Выберите файлы для обработки и конечную папку!')
    else:
        messagebox.showinfo('ЦОПП Бурятия','Создание отчета успешно завершено!')


if __name__ == '__main__':
    window = Tk()
    window.title('ЦОПП Бурятия')
    window.geometry('700x660')
    window.resizable(False, False)


    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку обработки данных для Приложения 6
    tab_report = ttk.Frame(tab_control)
    tab_control.add(tab_report, text='Скрипт №1')
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

    window.mainloop()