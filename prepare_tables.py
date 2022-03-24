import pandas as pd
import os


path = 'resources/MO/'
clean_path = 'resources/clean/MO/'

for file in os.listdir(path):
    print(file)
    name_file = file.split('.')[0]

    temp_df = pd.read_excel(f'{path}/{file}', sheet_name=0,
                            dtype={'ИНН': str, 'Телефон': str, 'Номер паспорта': str,
                                   'Серия паспорта': str,'Дата выдачи паспорта':str,'Дата рождения':str})
    temp_df.fillna('НЕ ЗАПОЛНЕНО!!!', inplace=True)
    # Очищаем от пробелов
    temp_df['Номер паспорта'] = temp_df['Номер паспорта'].apply(lambda x:x.replace(' ',''))
    temp_df['Серия паспорта'] = temp_df['Серия паспорта'].apply(lambda x: x.replace(' ', ''))
    temp_df['Дата рождения'] = temp_df['Дата рождения'].apply(lambda x: x.replace(',', '.').strip())

    temp_df.to_excel(f'{clean_path}{name_file} Очищеннный.xlsx',index=False)

