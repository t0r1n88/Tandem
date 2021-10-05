import pandas as pd

# df_abitur = pd.read_excel('resources/list_abitur_test.xlsx')
# df_prikaz = pd.read_excel('resources/list_prikaz_test.xlsx')

df_abitur = pd.read_csv('resources/csv_abitur.csv',encoding='cp1251')
df_prikaz = pd.read_csv('resources/csv_prikaz.csv',encoding='cp1251')


set_abitur = set(df_abitur['ФИО'])
set_prikaz = set(df_prikaz['ФИО'])
#
print((len(set_abitur)))
print((len(set_prikaz)))

lst_diff = []

for fio in set_prikaz:
    if fio not in set_abitur:
        lst_diff.append(fio)

print(len(lst_diff))
print(lst_diff)
# #
# # #
# # set_abitur = {'Алексеев Александр Сергеевич','Агеев Роман Алексеевич',  'Абидуев Александр Валерьевич'}
# # set_prikaz = {'Агеев Роман Алексеевич', 'Алексеев Александр Сергеевич'}
# #
# tr = set_prikaz-set_abitur
#
# print(len(tr))
# print(tr)
#

# print(df_abitur.head())
#
# print('*******')
#
# print(df_prikaz.head())

