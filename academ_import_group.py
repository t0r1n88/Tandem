import pandas as pd
import os
import random

path = 'resources/data'

for file in os.listdir(path):
    academ_df = pd.DataFrame()
    df = pd.read_excel(f'{path}/{file}')
    academ_df['Имя'] = df['Имя']
    academ_df['Отчество'] = df['Отчество']
    academ_df['Фамилия'] = df['Фамилия']
    academ_df['email'] = df['e-mail']
    academ_df['Пароль'] = random.randint(111111,999999)
    academ_df['Страна'] = 'RU'
    academ_df['Тип'] = 'ST'

    name_file = file.split('.')[0]
    academ_df.to_csv(f'resources/academ_group/{name_file}.csv',index=False,encoding='cp1251',sep =';')
