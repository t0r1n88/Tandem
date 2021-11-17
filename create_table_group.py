import pandas as pd

# Отражение максимального количества колонок в пайчарме
pd.set_option('display.max_columns', None)
base_df = pd.read_excel('base_to_import.xlsx')

edu_name_df = pd.read_excel('resources/edu_ou_t.xlsx')
# Отбираем колонку с кодом специальности и id
short_edu_df = edu_name_df[['short_title_p', 'id']]
print(short_edu_df.head())

# Отбираем нужные колонки
group_df = base_df[['Группа', 'Код специальности', 'ср. балл атт.', 'Год приема в БРИТ', 'Текущий курс']]
# Группируем по 2 столбцам
agg_group = group_df.groupby(['Группа', 'Код специальности', 'Год приема в БРИТ', 'Текущий курс'])['ср. балл атт.'].count()
# добавляем индекс
agg_group = agg_group.reset_index()
# Удаляем кололонку со средним баллом
agg_group.drop(['ср. балл атт.'], inplace=True, axis=1)
print(agg_group.shape)
agg_group.to_excel('resources/groups_codes.xlsx')

# Соединяем 2 датафрейма по коду специальности

itog_group = pd.merge(agg_group, short_edu_df, left_on='Код специальности', right_on='short_title_p')

itog_group.to_excel('resources/itog_group.xlsx', index=False)

group_t_df = pd.DataFrame()
group_t_df['id'] = range(itog_group.shape[0])
group_t_df['edu_ou_id'] = itog_group['id']
group_t_df['name_p'] = itog_group['Группа']
group_t_df['year_p'] = itog_group['Год приема в БРИТ']
group_t_df['course_p'] = itog_group['Текущий курс']

print(group_t_df.head())
group_t_df.to_excel('resources/group_t.xlsx',index=False)