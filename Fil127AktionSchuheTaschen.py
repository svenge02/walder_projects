import pandas as pd
from xlwings import view

df = pd.read_csv('G:\controlling\python\daten\Fil127.csv', encoding='cp1252',
                 sep=';', thousands="'")
print(df.shape)
print(df)
df['ART'] = df['ART'].apply(str)
# df['WG'] = df['ART'].astype(str).str[0]
df['WG'] = df['ART'].str[0]


df_neu = df[df.VIW <= df.VSW * 0.8]
view(df_neu)
print(df_neu.info())

view(df_neu.pivot_table(index=['DATE', 'BON', 'WG'], values='MEN', margins=True, fill_value=0, aggfunc='sum', margins_name = 'Total'))

