import tkinter
from create_report_priemka import processing_report_tandem # скрипт для обработки отчета по приемке
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
        processing_report_tandem(name_file_person,path_to_end_folder_report)


    except NameError:
        messagebox.showerror('Создание отчета приемной комиссии',
                             'Выберите файлы для обработки и конечную папку!')


"""
Функции для проверки наличия людей
"""




if __name__ == '__main__':
    window = Tk()
    window.title('ЦОПП Бурятия Создание отчета приемной комиссии ver 1.93')
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


    # Создаем кнопку Выбрать файл с данными персон
    lbl_person = Label(tab_report,
                      text='Чтобы получить файл нажмите\n'
                           'Абитуриенты-Отчеты-Выборка абитуриентов-Получить')
    lbl_person.grid(column=0, row=1, padx=10, pady=25)
    btn_choose_data_person = Button(tab_report, text='1) Выберите файл выборки'
                                                     , font=('Arial Bold', 20),
                                    command=select_file_data_person
                                    )
    btn_choose_data_person.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder = Button(tab_report, text='2) Выберите конечную папку', font=('Arial Bold', 20),
                                   command=select_end_folder
                                   )
    btn_choose_end_folder.grid(column=0, row=4, padx=10, pady=10)

    # Создаем кнопку обработки данных

    btn_proccessing_data = Button(tab_report, text='3) Создать отчет', font=('Arial Bold', 20),
                                  command=processing_report
                                  )
    btn_proccessing_data.grid(column=0, row=5, padx=10, pady=10)


    window.mainloop()