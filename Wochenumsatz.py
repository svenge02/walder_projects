# für Dataframes
import pandas as pd
# für die Integration in Excel
import xlwings as xw



wb = xw.Book(r'G:\controlling\python\Berichte\Wochenumsatz\Wochenumsatz2021_2.xls')
df = pd.read_csv(r'G:\controlling\python\daten\wochenumsatz.csv', encoding='cp1252',
                 sep=';', thousands="'")

df_dummy = pd.read_csv(r'G:\controlling\python\daten\wochenumsatz_dummy.csv', encoding='cp1252',
                 sep=';', thousands="'")



df_alle = df.append(df_dummy)

# Zeitraum (integer) in String umwandeln
df_alle['ZR'] = df['ZR'].apply(str)


# Angaben Jahr in Zeitraum entfernen und in Woche abspeichern

# df['woche'] = df['ZR'].str.replace("2021","")
df_alle['woche'] = df['ZR']

xw.sheets('apollon').range('a4').value = df_alle.pivot_table(index=['woche'], columns=['FIL'], values='UVKI', fill_value=0, aggfunc='sum')

print(df_alle)
print(df_alle.info())