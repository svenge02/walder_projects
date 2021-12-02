import pandas as pd
import xlwings as xw


df = pd.read_csv('G:\controlling\python\daten\Lager_20_02.csv', encoding='cp1252', sep=';')
df.drop(['WAG', 'SAISON', 'EMEN', 'UMEN', 'BMEN', 'EEKN', 'EVKS', 'UVKS','UEKN', 'BVKS', 'UVKI', 'BVKA', 'EANT', \
         'UANT', 'BANT', 'VERLUST', 'PLANVERLUST', 'VERLUSTPROZ','SSPANNE', 'ISPANNE', 'BSSPANNE', 'ABVW', 'ABVM' \
        ], axis = 1, inplace = True)

print(df)
print(df.info())

print('df')

df['BEKN'] = pd.to_numeric(df['BEKN'])


# neues DataFRame als Pivot-Tabelle ohne Totale darstellen und in Excel-Tabelle schreiben.
xw.view(df.pivot_table(index='FILIALE', columns=['Jahr','Monat'],values='BEKN', margins = True, fill_value=0, aggfunc='sum'))
