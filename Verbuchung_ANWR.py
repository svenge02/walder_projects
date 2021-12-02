import csv
import os
import pandas as pd
import datetime as dt
import xlwings as xw
import hashlib
import re


def datei_eingelesen():
    """Überprüft anhand des Hash-Wertes einer Datei, ob diese schon einmal verarbeitet wurde. Zuerst wird
    der Dateiname in einen Hash-Wert umgewandelt und mit einer Liste verglichen, falls er nicht in der Liste
    ist, wir er ebenfalls der Liste hinzugefügt."""
    with open(pfad + os.sep + datei, "rb") as f:
        file_hash = hashlib.md5()
        while chunk := f.read(8192):
            file_hash.update(chunk)
    hash = file_hash.hexdigest()
    print(hash_list)
    for entry in hash_list:
        if re.search(hash, str(entry)):
            return True
    hash_list.append(hash)
    return False

# def konto_holen(row):
#    if row in dic_kreditoren:
#        return dic_kreditoren[]
#    return df_kreditoren['Fibu-Konto'] == row


def mwst_code_bestimmen(row):
    if row:
        return ohne_mwst
    else:
        return mit_mwst


def datum_vergleichen(row):
    if letztes_buchungsdatum < row:
        return row
    else:
        return letztes_buchungsdatum


kopfzeile = []
daten = []
pfad = r'G:\Controlling\Python\ANWR-Verbuchung\csv_files'
pfad_export = r'G:\Controlling\Python\ANWR-Verbuchung\export'
pfad_system = r'G:\Controlling\Python\ANWR-Verbuchung\system'
pafd_backup = r'G:\Controlling\Python\ANWR-Verbuchung\backup'

letztes_buchungsdatum = dt.datetime(2021, 8, 1)
mit_mwst = 140
ohne_mwst = 100
wb = xw.Book(pfad_export + os.sep + "Datenexport_Fibusync.xlsx")

# alte komplett Datei löschen, falls vorhanden
file_test = pfad + os.sep + 'komplett.csv'
try:
    os.remove(file_test)
except OSError as e:
    print(e)
else:
    print("Alte Datei wurde gelöscht")

df_hash_list = pd.read_csv(pfad_system + os.sep + "hashes.csv")
hash_list = df_hash_list.values.tolist()

dateien = os.listdir(pfad)


for datei in dateien:
    if datei.lower().endswith('.csv'):
        if datei_eingelesen() == False:
            with open(pfad + os.sep + datei, 'r') as ausgleich_daten:
                reader = csv.reader(ausgleich_daten, delimiter=";")
                kopfzeile = next(reader)
                daten.extend([zeile for zeile in reader])

            with open(pfad + os.sep + 'komplett.csv', 'w', newline='') as ergebnis_daten:
                writer = csv.writer(ergebnis_daten, delimiter=";")
                writer.writerow(kopfzeile)
                writer.writerows(daten)


df_neue_hash_liste = pd.DataFrame(data={"col1": hash_list})
df_neue_hash_liste.to_csv(pfad_system + os.sep + "hashes.csv",
                          header=False,
                          index=False,
                          doublequote=False,
                          quoting=None,
                          escapechar=None,
                          sep=';')

komplett = (pfad + os.sep + 'komplett.csv')


df_alle = pd.read_csv(komplett,
                      encoding='cp1252',
                      sep=';',
                      usecols=[1, 4, 5, 6, 9, 12,  13, 15, 18, 23, 26],
                      )
df_alle.columns = ['Ausgleichsdatum', 'Positionstext', 'Kredi-Nr', 'Buchungstext', 'Kontonummer', 'Belegart',
                   'Belegdatum', 'Belegnummer', 'Rechnung', 'Skonto', 'Überweisung']

# df_alle.columns = ['Ausgleichsdatum', 'Positionstext', 'Kredi-Nr', 'Buchungstext', 'Kontonummer', 'Belegart',
#                   'Belegdatum', 'Belegnummer', 'Rechnung', 'Skonto', 'Überweisung']

print(df_alle.info())
df_kreditoren = pd.read_csv(pfad_system + os.sep + "Kreditoren.csv",
                            encoding='cp1252',
                            usecols=[0, 2],
                            sep=';'
                            )

# Belegdatum von object in Datum umwandeln
df_alle['Ausgleichsdatum'] = pd.to_datetime(df_alle['Ausgleichsdatum'])
df_alle['Belegdatum'] = pd.to_datetime(df_alle['Belegdatum'], format='%d.%m.%Y')


# Belegdatum auf letztes Buchungsdatum ändern, falls das Belegdatum älter als letztes Buchungsdatum ist
df_alle['Belegdatum'] = df_alle['Belegdatum'].apply(datum_vergleichen)

df_alle['mwst_code'] = df_alle['Skonto'].isnull().apply(mwst_code_bestimmen)
df_alle['Rechnung'] = df_alle['Rechnung'].str.replace(",", ".").astype(float)
df_alle['Skonto'] = df_alle['Skonto'].str.replace(",", ".").astype(float)
df_alle['Überweisung'] = df_alle['Überweisung'].str.replace(",", ".").astype(float)
df_alle['Buchungstext'] = df_alle['Buchungstext'].str[0:11] + " " + df_alle['Belegnummer']
print(df_alle)
pattern = '^2021-..'
mask_position = (df_alle['Positionstext'].str.match(pattern) == True)
mask_belegart = (df_alle['Belegart'].isnull() == True)

df_loeschen = df_alle.loc[mask_position & mask_belegart]

zum_loeschen = df_loeschen.index.tolist()

print(zum_loeschen)

df_alle.drop(zum_loeschen, inplace = True, axis = 0)
print(df_alle)
df_mit_Kredi = df_alle.merge(df_kreditoren,
                           how="left",
                           left_on="Kredi-Nr",
                           right_on="Nummer",
                           validate="many_to_one",
                           indicator=True)

df_mit_Kredi.drop(['Positionstext', 'Kredi-Nr', 'Belegart', 'Nummer', '_merge'], inplace=True, axis=1)
print(df_mit_Kredi)
xw.sheets['kredi'].range('c5').value = df_mit_Kredi
df_chf = df_mit_Kredi[df_mit_Kredi['Kontonummer'] == 1790046937]
df_eur = df_mit_Kredi[df_mit_Kredi['Kontonummer'] == 1790047075]

df_chf.drop(['Kontonummer'], inplace=True, axis=1)
df_eur.drop(['Kontonummer'], inplace=True, axis=1)



xw.sheets['chf'].range('c5').value = df_chf
xw.sheets['eur'].range('c5').value = df_eur

df_chf.to_csv(pfad_export + os.sep + "chf_Fibu_Sync.csv",
              header=True,
              columns=['Ausgleichsdatum', 'Belegdatum'],
              index=False,
              sep=';')
