# Versandretouren in den Filialen
# Erstellt eine Liste in einer Excel-Tabelle mit den Retouren, die in die Filiale zurückgebracht wurden
# wird neu jetzt monatlich erstellt und verteilt.
#
# SG/27.04.2021

import pandas as pd
import xlwings as xw

# Öffnen der benötigten Tabelle
wb = xw.Book(r'G:\controlling\python\Berichte\2021_Versandretouren_in_Filialen.xlsx')

# Einlesen des CSV-Files das in Apollon erstellt wurde Titel= Versandretouren 095 neu
df = pd.read_csv('G:\controlling\python\daten\Versand_Retouren_Filialen_2021.csv', encoding='cp1252', sep=';')


# DATE in DateTime-Objekt umwandeln
df['DATE'] = pd.to_datetime(df['DATE'], dayfirst = True)

# Tabelle um Jahr und Monatsfeld ergänzen (wird für die Pivot-Tabelle benötigt.
df['Year'] = pd.DatetimeIndex(df['DATE']).year
df['Month'] = pd.DatetimeIndex(df['DATE']).month

# Wert von negativ zu positiv umwandeln, damit die MA in der Fil nicht verwirrt werden.
df['VIW'] = df['VIW'] * -1

# neues DataFrame erstellen, mit nur Retouren der Mitarbeiternummer 095
df_neu = df[df.VKN == 95]

# neues DataFRame als Pivot-Tabelle ohne Totale darstellen und in Excel-Tabelle schreiben.
xw.sheets('2021').range('B5').value = df_neu.pivot_table(index='FIL', columns=['Year','Month'],values='VIW', fill_value=0, aggfunc='sum')
