import pandas as pd
import os
# Для того чтобы не показывало предупреждение UserWarning: Data Validation extension is not supported and will be removed
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
pd.set_option('display.max_columns', None)
"""
Скрипт для заполнения пустых email и очистки некоторых полей от пробельных символов.
"""
path = 'resources/data/'

# Переменные для почты
adress = 'brit_student'
end_mail = '@mail.ru'
count = 1

# Перебираем
for file in os.listdir(path):
    # print(file)
    df = pd.read_excel(f'{path}/{file}',dtype={'Серия паспорта':str,'Номер паспорта':str,'Номер документа о образовании':str})
    for row in df.itertuples():
        # print(row)
        # Это поле почты row[18]
        # Поле аттестата  row[25]
    #     # Сначала Проверяем на пустое значение
        if pd.isnull(row[18]):
            df.loc[row[0],'e-mail'] = f'{adress}{count}{end_mail}'
            count += 1
        # Только после этого когда мы избавились от nan очищаем полe почты от пробельных символов
        email = df.loc[row[0],'e-mail']
        # Проверяем наличие аттестата.Если пусто то пропускаем, так как это проверяется на следующем этапе
        if pd.isnull(row[25]):
            continue
        # Тоже самое с серией и номером паспорта
        if pd.isnull(row[4]) or pd.isnull(row[5]):
            continue

        df.loc[row[0], 'e-mail'] = email.strip().replace(' ','')
        df.loc[row[0],'Номер документа о образовании'] = row[25].strip().replace(' ','')
        df.loc[row[0], 'Серия паспорта'] = row[4].strip().replace(' ', '')
        df.loc[row[0], 'Номер паспорта'] = row[5].strip().replace(' ', '')
    # Удаляем лишнюю колонку
    # df.drop('Unnamed: 0',axis=1,inplace=True)


    #Сохраняем файл
    df.to_excel(f'resources/fill_email/{file}',index=False)
with open('Число пустых email.txt','w',encoding='utf-8') as f:
    f.write(str(count))

