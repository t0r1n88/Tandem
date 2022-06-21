import pandas as pd
import openpyxl
import numpy as np
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
import time
import datetime
from datetime import date

name_file_abiturs = 'data/abitur.xlsx'
name_file_person = 'data/person.xlsx'
path_to_end_folder_report = 'data'

df_abitur = pd.read_excel(name_file_abiturs,skiprows=3,usecols=['Абитуриент','Доп. статус','Состояние'])
df_person = pd.read_excel(name_file_person,sheet_name='Абитуриенты',skiprows=8,usecols=['ФИО','Нуждается в общежитии','Формирующее подр.','Направление, специальность, профессия','Сдан оригинал'])

wb = openpyxl.Workbook()
# Переименовываем лист
sheet = wb['Sheet']
sheet.title = 'Отчет'

# Соединяем оба датафрейма

cross_df = df_person.merge(df_abitur,how='inner',left_on='ФИО',right_on='Абитуриент')



# Преобразовываем да-нет в 1 или 0 для подсчетов
cross_df['Нуждается в общежитии'] =cross_df['Нуждается в общежитии'].apply(lambda x:0 if x =='нет' else 1)
cross_df['Сдан оригинал'] =cross_df['Сдан оригинал'].apply(lambda x:0 if x =='нет' else 1)
cross_df['Состояние'] =cross_df['Состояние'].apply(lambda x:1 if x =='Забрал документы' else 0)


# заменяем нан на пустые строки чтобы произвести поиск слова сирота;
cross_df['Доп. статус'].fillna('',inplace=True)
cross_df['Доп. статус'] = cross_df['Доп. статус'].apply(lambda x:1 if 'Сирота;' in x else 0)

cross_df['for_counting'] = 1

svod_df = pd.DataFrame.pivot_table(cross_df,index=['Формирующее подр.','Направление, специальность, профессия'],
                     values=['for_counting','Состояние','Сдан оригинал','Доп. статус','Нуждается в общежитии'],
                     aggfunc='sum')

svod_df.columns = ['Сдали всего','Сирот чел.','Нуждается в общежитии чел.','Сдано оригиналов','Забрали заявления']

svod_df['Итого'] = svod_df['Сдали всего'] - svod_df['Забрали заявления']

# Меняем местами столбцы
out_df = svod_df.reindex(columns=['Сдали всего','Забрали заявления','Итого','Сдано оригиналов','Сирот чел.','Нуждается в общежитии чел.'])

# разворачиваем столбец в строку
sum_row=out_df.sum(axis=0).to_frame().transpose()

# Добавляем колонки чтобы сделать из них мультинидекс .Ужасно решение но что есть то есть
sum_row['1'] ='Всего'
sum_row['2'] = ''

# Делем мультинидекс и объединяем датафреймы
sum_row.set_index(['1','2'],inplace=True)
all_out_df = pd.concat([out_df,sum_row],axis=0)



#Преобразовываем мультинидекс в колонки
finish_df=all_out_df.reset_index()

for r in dataframe_to_rows(finish_df,index =False,header=True):
    if len(r) != 1:
        wb['Отчет'].append(r)

# # Настраиваем выходной файл
wb['Отчет'].column_dimensions['A'].width =30
wb['Отчет'].column_dimensions['B'].width =50
wb['Отчет']['B2'].alignment = Alignment(wrap_text=True)
wb['Отчет'].column_dimensions['C'].width =20
wb['Отчет'].column_dimensions['D'].width =20
wb['Отчет'].column_dimensions['F'].width =20
wb['Отчет'].column_dimensions['G'].width =20
wb['Отчет'].column_dimensions['H'].width =30
wb['Отчет']['H1'].alignment = Alignment(wrap_text=True)

 # Получаем текущее время для того чтобы использовать в названии
t = time.localtime()
current_time = time.strftime('%H_%M_%S', t)
# Сохраняем итоговый файл
wb.save(f'{path_to_end_folder_report}/Ежедневный отчет приемной комиссии ГБПОУ БРИТ {current_time}.xlsx')







