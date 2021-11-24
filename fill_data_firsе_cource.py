import pandas as pd
import os

# Считываем данные
person_df = pd.read_excel('resources/person_t.xlsx')
student_df = pd.read_excel('resources/student_t.xlsx')
base_df = pd.read_excel('base_to_import.xlsx')

# Получаем датафрейм с инн,снилс и пр.
base_person_part_df = base_df[['Место рождения','СНИЛС','ИНН','Телефон','e-mail']]
print(base_person_part_df.head())


