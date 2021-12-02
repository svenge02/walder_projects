# Versandretouren in den Filialen
# Erstellt eine Liste in einer Excel-Tabelle mit den Retouren, die in die Filiale zurückgebracht wurden
# wird neu jetzt monatlich erstellt und verteilt.
#
# SG/27.04.2021

import pandas as pd
import xlwings as xw
import numpy as np




# Öffnen der benötigten Tabelle
wb = xw.Book(r'G:\controlling\python\Berichte\2021_Versand_095.xlsx')

# Einlesen des CSV-Files das in Apollon erstellt wurde Titel= Versandretouren 095 neu
df = pd.read_csv('G:\controlling\python\daten\Versand_095.csv', encoding='cp1252', sep=';')


# DATE in DateTime-Objekt umwandeln
df['DATE'] = pd.to_datetime(df['DATE'], dayfirst = True)

# Tabelle um Jahr und Monatsfeld ergänzen (wird für die Pivot-Tabelle benötigt.
df['Jahr'] = pd.DatetimeIndex(df['DATE']).year
df['Monat'] = pd.DatetimeIndex(df['DATE']).month


df['Umsatz'] = np.where(df['VIW'] >= 0, df['VIW'], 0)
df['Retouren'] = np.where(df['VIW'] < 0, df['VIW'], 0)



# neues DataFRame als Pivot-Tabelle ohne Totale darstellen und in Excel-Tabelle schreiben.
xw.sheets('2021').range('B22').value = df.pivot_table(index=['Jahr', 'Monat'], values=['Umsatz', 'Retouren'], fill_value=0, aggfunc='sum')