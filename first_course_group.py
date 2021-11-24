import pandas as pd

df = pd.read_excel('resources/Для групп.xlsx')
group_df = pd.read_excel('resources/group_t.xlsx')

out_group_df = pd.merge(df,group_df,how='outer',left_on='Группа',right_on='name_p')

out_group_df.to_excel('Проверка.xlsx',index=False)
out_group_df.drop(['edu_ou_id','name_p','year_p','course_p'],axis=1,inplace=True)
out_group_df.rename(columns={'id_x':'ID_P','id_y':'ID_GROUP'},inplace=True)

out_group_df.to_excel('ФИО + группа.xlsx',index=False)


# Заполняем группы для студентов

student_df = pd.read_excel('resources/student_t.xlsx')

merged_df = pd.merge(student_df,out_group_df,left_on='person_id',right_on='ID_P')

# Присваиваем айди группы и удаляем лишние столбы
merged_df['group_id'] = merged_df['ID_GROUP']

merged_df.drop(['ID_P','ФИО','Группа','ID_GROUP'],axis=1,inplace=True)

merged_df.to_excel('Студенты + группы.xlsx',index=False)
