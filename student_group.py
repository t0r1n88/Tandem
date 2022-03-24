import pandas as pd
# Отражение максимального количества колонок в пайчарме
pd.set_option('display.max_columns', None)

# Создали базовый датафрейм
df = pd.read_excel('resources/Исправленное Для групп 2-3 курс.xlsx')
group_df = pd.read_excel('resources/group_t.xlsx')

out_group_df = pd.merge(df,group_df,how='outer',left_on='Группа',right_on='name_p')

out_group_df.to_excel('Проверка.xlsx',index=False)
# out_group_df.drop(['edu_ou_id','name_p','year_p','course_p'],axis=1,inplace=True)
out_group_df.rename(columns={'id_x':'ID_P','id_y':'ID_GROUP'},inplace=True)

out_group_df.to_excel('ФИО + группа.xlsx',index=False)

# Загружаем очищенную от первокурсников таблицу
df23 = pd.read_excel('resources/база для заполнения student_p 2-3 курс.xlsx')


import_df = pd.DataFrame()

import_df['id'] = range(1,df23.shape[0] + 1)
import_df['person_id'] = df23['ID_P']
import_df['number_p'] = ''
import_df['entrance_p'] = df23['year_p']
import_df['group_id'] = df23['ID_GROUP']
import_df['edu_ou_id'] = df23['edu_ou_id']
import_df['course_p'] = df23['course_p']
import_df['compensation_p'] = '1 Госбюджетное место'
import_df['category_p'] = '1 Студент'
import_df['bookNumber_p'] = ''
import_df['status_p'] = '1 Активный'
import_df['archival_p'] = '0'
import_df['finishYear_p'] = ''
import_df['personalFileNumber_p'] = ''
import_df['specialPurposeRecruit_p'] = '0'
import_df['edu_document_id'] = ''
import_df['edu_document_original'] = '1'
import_df['contractNumber'] = ''
import_df['contractDate'] = ''
import_df['userCode'] = ''

print(import_df.tail())
print(import_df.shape)




# Заполняем группы для студентов

student_df = pd.read_excel('resources/student_t.xlsx')
print(student_df.shape)

# нам нужно не мержить в данном случае а добавить датафрейм через append

student_df = student_df.append(import_df)
print(student_df.shape)
student_df.to_excel('Студенты 2-3 курс 1 попытка.xlsx',index=False)





# ДЛЯ ПЕРВОГО КУРСА
# merged_df = pd.merge(student_df,out_group_df,how='outer',left_on='person_id',right_on='ID_P')

# Создаем дф для добавления студентов



# # Присваиваем айди группы и удаляем лишние столбы
# merged_df['group_id'] = merged_df['ID_GROUP']
#
# merged_df.drop(['ID_P','ФИО','Группа','ID_GROUP'],axis=1,inplace=True)
#
# merged_df.to_excel('Студенты + группы.xlsx',index=False)
