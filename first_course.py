import pandas as pd
import os
# Отражение максимального количества колонок в пайчарме
pd.set_option('display.max_columns', None)
# Считываем базовый датафрейм с первокурсниками
base_df = pd.read_excel('resources/data_first_course.xlsx',dtype={'Дата выдачи паспорта':str,'Дата рождения':str})
# Создаем столбец ФИО
base_df['ФИО'] = base_df['Фамилия']+' ' + base_df['Имя']+' ' + base_df['Отчество']

# Создаем датафрейм с нужными столбцами
df = base_df[['ФИО','ИНН','СНИЛС','e-mail','Телефон','Группа']]

# Считываем главный датафрейм
person_df = pd.read_excel('resources/person_t.xlsx')
person_df['ФИО'] = person_df['identity_lastName'] +' '+ person_df['identity_firstName'] + ' '+ person_df['identity_middleName']

itog_df = pd.merge(person_df, df, how='outer',left_on='ФИО', right_on='ФИО')

# itog_df.to_excel('Первый курс пробв.xlsx')

itog_df['inn_number'] = itog_df['ИНН']
itog_df['snils_number'] = itog_df['СНИЛС']
itog_df['email'] = itog_df['e-mail']
itog_df['phoneMobile'] = itog_df['Телефон']

# Сохраняем датафрейм чтобы соединить с группами
group_df = itog_df[['id','ФИО','Группа']]

group_df.to_excel('Для групп.xlsx',index=False)
# itog_df.drop(['ФИО','ИНН','СНИЛС','e-mail','Телефон','Группа'],inplace=True,axis=1)




# Сохраняем в итоги
# itog_df.to_excel('Первый курс persont_t.xlsx',index=False)



# missed_df = pd.read_excel('resources/missed_first_course.xlsx')
# missed_out_df = pd.merge(missed_df,df,left_on='e-mail',right_on='e-mail')
# missed_out_df.to_excel('ФИО пропущенных.xlsx',index=False)
