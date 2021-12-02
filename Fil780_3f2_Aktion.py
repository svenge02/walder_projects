import pandas as pd
from xlwings import view

df = pd.read_csv('G:\controlling\python\daten\Fil780.csv', encoding='cp1252', sep=';')
print(df.shape)

# df['UVKI'] = pd.to_numeric(df['UVKI'])


df_neu = df[df.VIW == 0]
print(df_neu)


view(df_neu.pivot_table(index=['WG1', 'WG2'], values='MEN', margins=True, fill_value=0, aggfunc='sum', margins_name = 'Total'))

