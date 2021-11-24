import pandas as pd
import os
# Отражение максимального количества колонок в пайчарме
pd.set_option('display.max_columns', None)

# Базовый датафрейм с исходными данными
base_df = pd.read_excel('resources/2-3 course.xlsx')

# Датафрейм для экспорта в эксель
person_df = pd.read_excel('resources/person_t.xlsx')

# Промежуточный датафрейм
df = pd.DataFrame()

# Заполняем датафрейм
df['id'] = range(1,base_df.shape[0])
df['identity_type_p'] = '1 Паспорт гражданина Российской Федерации'
df['identity_seria'] = base_df['Серия паспорта']
df['identity_number'] = base_df['Номер паспорта']
df['identity_firstName'] = base_df['Имя']
df['identity_lastName'] = base_df['Фамилия']
df['identity_middleName'] = base_df['Отчество']
df['identity_birthDate'] = base_df['Дата рождения']
df['identity_birthPlace'] = base_df['Место рождения']
df['identity_seria'] = base_df['Серия паспорта']





print(df.head())
