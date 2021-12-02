import pandas as pd
from xlwings import view

df = pd.read_csv('G:\controlling\python\daten\zusatzartikel.csv', encoding='cp1252', sep=';', thousands="'")
print(df.shape)

# df['UVKI'] = pd.to_numeric(df['UVKI'])

# df.drop(df.WAGBEZ['Reparaturen'])

df_neu = df[df.WAGBEZ != 'Reparaturen']
df_neu = df_neu[df_neu.WAGBEZ != 'ProspektRabatt']
print(df_neu)


view(df_neu.pivot_table(index='FIL', columns=['WAGBEZ'], values='UVKI', margins=True, fill_value=0, aggfunc='sum'))
print(df_neu.info())