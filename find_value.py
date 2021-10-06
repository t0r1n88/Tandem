import pandas as pd

df = pd.read_excel('resources/lst_abitur.xlsx')
df['Метка'] =df['Адрес регистрации'].apply(lambda x:'Кабанск' if 'кабанск' in x.lower() else x)
df.to_excel('Kabansk.xlsx',index=False)