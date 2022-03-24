import pandas as pd
import os
# Отражение максимального количества колонок в пайчарме
pd.set_option('display.max_columns', None)

# Базовый датафрейм с исходными данными
base_df = pd.read_excel('resources/2-3 course.xlsx',dtype={'ИНН': str, 'Телефон': str, 'Номер паспорта': str,
                                       'Серия паспорта': str})


# Датафрейм для экспорта в эксель
person_df = pd.read_excel('resources/person_t.xlsx',dtype={'inn_number':str,'phoneMobile':str,'identity_seria':str,'identity_number':str})

# Промежуточный датафрейм
df = pd.DataFrame()

# Заполняем датафрейм
df['id'] = range(1,base_df.shape[0] + 1)
df['identity_type_p'] = '1 Паспорт гражданина Российской Федерации'
df['identity_seria'] = base_df['Серия паспорта']
df['identity_number'] = base_df['Номер паспорта']
df['identity_firstName'] = base_df['Имя']
df['identity_lastName'] = base_df['Фамилия']
df['identity_middleName'] = base_df['Отчество']
df['identity_birthDate'] = base_df['Дата рождения']
df['identity_birthPlace'] = base_df['Место рождения']
df['identity_sex_p'] = base_df['Пол']
# Заменим значения на принятые в тандеме
df['identity_sex_p'] = df['identity_sex_p'].apply(lambda x:'2 Женский' if x =='Женский' else '1 Мужской')
df['identity_date_p'] = base_df['Дата выдачи паспорта']
df['identity_code_p'] = base_df['Код подразделения']
df['identity_place_p'] = base_df['Кем выдан паспорт']
df['identity_citizenship_p'] = '0 Россия'
df['family_status'] = ''
df['child_count'] = ''
df['pension_type'] = ''
df['pension_issuance_date'] = ''
df['flat_presence'] = ''
df['service_length_years'] = 0
df['service_length_months'] = 0
df['service_length_days'] = 0

df['address_reg_id'] = ''
df['address_fact_id'] = ''
df['inn_number'] = base_df['ИНН']
df['snils_number'] = base_df['СНИЛС']
df['email'] = base_df['e-mail']
df['phoneDefault'] = ''
df['phoneMobile'] = base_df['Телефон']
df['phoneFact'] = ''
df['phoneReg'] = ''
df['phoneRegTemp'] = ''
df['phoneRelatives'] = ''
df['phoneWork'] = ''
df['login'] = ''
print(person_df.shape)
print(df.shape)


itog_df = person_df.append(df,ignore_index=True)
print(itog_df.shape)
# itog_df.to_excel('Персоны 2-3 курс.xlsx',index=False)

df.to_excel('Персоны 2-3 курс.xlsx',index=False)









