import pandas as pd
import datetime
df = pd.read_excel('date.xlsx')


# неужели я понял как это сделать. Сначала привести к датетайм,  потом к нужному формату через dt.strftime
df['Дата выдачи паспорта'] = pd.to_datetime(df['Дата выдачи паспорта'],format='%Y-%m-%d')
df['Дата выдачи паспорта'] = df['Дата выдачи паспорта'].dt.strftime('%d.%m.%Y')