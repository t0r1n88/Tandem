import pandas as pd
import os
# Для того чтобы не показывало предупреждение UserWarning: Data Validation extension is not supported and will be removed
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
pd.set_option('display.max_columns', None)
"""
Скрипт для заполнения пустых email 
"""
path = 'resources/data/'

# Переменные для почты
adress = 'brit_student'
end_mail = '@mail.ru'
count = 1

# Перебираем
for file in os.listdir(path):
    df = pd.read_excel(f'{path}/{file}')
    for row in df.itertuples():
        # print(row)
        email = row[18]
    #     # Проверяем на пустое значение
        if pd.isnull(email):
            df.loc[row[0],'e-mail'] = f'{adress}{count}{end_mail}'
            count += 1

    df.to_excel(f'resources/fill_email/{file}')
with open('Число пустых email.txt','w',encoding='utf-8') as f:
    f.write(str(count))

