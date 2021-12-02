"""
Auswertung die für die eingelösten Geschenkgutscheine verwendet wird.
SG/25.06.2021
"""

import pandas as pd
import xlwings as xw
import numpy as np

week = [9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25,
        27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43,
        44, 45, 46]

# Gutscheine, die im Versand eingelöst wurden, erscheinen mit Fil und verschobenen Feldern
filialliste = [6]

# Öffnen der benötigten Tabelle
wb = xw.Book(r'G:\controlling\python\Berichte\2021_Gutscheine_HW.xlsx')

col = ['Filiale', 'Woche', 'Rec', 'Amount', 'LotId', 'LotName']

# Einlssen des CSV-Files das in der Gutscheinverwaltung erstellt und
# via PrimalSQL in ein CSV-File importiert wurde
df_alle = pd.read_csv(r'G:\controlling\python\daten\Gutscheine\waldermailing.csv',
                      encoding='cp1252',
                      sep=','
                      )
# Wird gebraucht, damit die Reihenfolge bei den Gutscheinen im Versand stimmt
df_dummy = pd.read_csv(r'G:\controlling\python\daten\Gutscheine\waldermailing_dummy.csv',
                       encoding='cp1252',
                       sep=';'
                       )
df_dummy.columns = col
df_alle.columns = col

df_versand = df_alle.loc[df_alle['Filiale'].isin(filialliste)]
df_gut = df_alle.loc[~df_alle['Filiale'].isin(filialliste)]

# dummy wid angehängt, damit in der Excel-Tabelle die Werte in der richtigen Reihe stehen
df_versand_alle = df_versand.append(df_dummy)

df_versand_alle["Amount_neu"] = df_versand.loc[:, ["Woche"]] / -100

df_gut['Amount_neu'] = np.where(df_gut.loc[:, ["Amount"]] == '-1000', 10.00, 20.00)
# df_gut['Woche'] = df_gut['Woche']
print(df_gut)

df_select = df_gut.loc[df_gut['Woche'].isin(week)]

xw.sheets('2021').range('B4').value = df_select.pivot_table(index=['Woche'],
                                                            columns=['LotId'],
                                                            values=['Amount_neu'], fill_value=0,
                                                            aggfunc='sum'
                                                            )

xw.sheets('2021').range('B26').value = df_select.pivot_table(index=['Filiale'],
                                                             columns=['LotId'],
                                                             values=['Amount_neu'], fill_value=0,
                                                             aggfunc='sum'
                                                             )

xw.sheets('2021').range('B20').value = df_versand_alle.pivot_table(index=['Filiale'],
                                                                   columns=['Rec'],
                                                                   values=['Amount_neu'], fill_value=0,
                                                                   aggfunc='sum'
                                                                   )