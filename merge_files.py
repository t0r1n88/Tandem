import pandas as pd
import os

# Путь к файлам котороые нужно соединить
path = 'resources/MO/'
# Базовый файл куда будут добавлятся данные
base_df = pd.read_excel('resources/base.xlsx')

# Обработка файлов
# Перебираем файлы в указанной директории. Создаем датафрейм из указанного листа
for file in os.listdir(path):
    temp_df = pd.read_excel(f'{path}/{file}', sheet_name='Список')
    base_df = base_df.append(temp_df)
# Вставляем столбец после ФИО, что логично
base_df.insert(3,'Наименование документа','Паспорт гражданина Российской Федерации')
# base_df['Наименование'] = 'Паспорт гражданина Российской Федерации'
print(base_df.shape)
print(base_df.head())
base_df.to_excel('base_to_import.xlsx',index=False)
