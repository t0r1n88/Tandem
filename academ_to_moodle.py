import pandas as pd
import os
# Отражение максимального количества колонок в пайчарме
pd.set_option('display.max_columns', None)
base_academ_df = pd.read_csv('resources/example_academ.csv',encoding='cp1251',sep=';')
moodle_df = pd.read_csv('resources/example_moodle.csv',encoding='cp1251',sep=';')

path = 'resources/academ_group_first_course/'
for file in os.listdir(path):

    print(file)
    name_group = file.split('.')[0]
    # Создаем временный датафрейм
    temp_df = pd.DataFrame()
    df = pd.read_csv(f'{path}/{file}',encoding='cp1251',sep=';')
    # Собираем наш  датафрейм
    temp_df['username'] = df['email'].apply(lambda x:x.lower())
    temp_df['password'] = df['Пароль']
    temp_df['firstname'] = df['Имя']
    temp_df['lastname'] = df['Фамилия']
    temp_df['email'] = df['email'].apply(lambda x:x.lower())
    temp_df['cohort1'] = name_group


    moodle_df = pd.concat([moodle_df,temp_df],ignore_index=True)

    group_df = temp_df.copy()
    group_df.insert(0,'ФИО','')
    group_df['ФИО'] = group_df['lastname'] + ' '+ group_df['firstname']
    group_df.rename(columns={'password':'Пароль','cohort1':'Группа'},inplace=True)
    group_df.drop(['firstname','username','lastname'],axis=1,inplace=True)
    group_df.to_excel(f'{name_group}.xlsx',index=False)
moodle_df.to_csv('firts_course2021.csv',encoding='cp1251',index=False,sep=';')

