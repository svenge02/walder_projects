# Dieses Programm erstellt eine Art Excel-Tabelle, die für die Abgrenzung
# der Wareneingänge benötigt wird.

import pandas as pd
import xlwings as xw

verw_spalten = (0, 1, 2, 8)

df = pd.read_csv(r'G:\controlling\python\daten\wareneingang\wareneingaenge.csv', encoding='cp1252',
                 usecols=verw_spalten, sep=';', thousands="'" )




df['DATUM'] = pd.to_datetime(df['DATUM'], dayfirst = True)
df['Jahr'] = pd.DatetimeIndex(df['DATUM']).year
df['Monat'] = pd.DatetimeIndex(df['DATUM']).month

df['Lieferant'] = df['LFNAME'] + ' ' + df['LFN'].apply(str)

xw.view(df.pivot_table(index=['Lieferant'], columns=['Jahr', 'Monat'], values='EKN'))
xw.view(df.pivot_table(index=['Lieferant', 'DATUM'], columns=['Jahr', 'Monat'], values='EKN'))



print(df.info())
print(df)