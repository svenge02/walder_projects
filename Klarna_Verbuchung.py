"""
Dieses Skript bereitet die Zahlen auf, damit sie über Fibusync verbucht werden können
SG/26.11.2021
"""

import pandas as pd
import xlwings as xw
import csv
import os
import datetime as dt
import shutil
import numpy as np

def datum_vergleichen(row):
    """überprüft ob, das Belegdatum älter als das zulässige Buchungsdatum ist und ändert dieses."""
    if letztes_buchungsdatum < row:
        return row
    else:
        return letztes_buchungsdatum

def alte_dateien_loeschen():
    """ falls vorhanden, wird die datei komplett.csv gelöscht"""
    file_test = pfad + os.sep + 'komplett.csv'
    try:
        os.remove(file_test)
    except OSError as e:
        print(e)
    else:
        print("Alte Datei wurde gelöscht")

# Definition verschiedener Variabeln und benötigter Pfade

kopfzeile = []
daten = []
letztes_buchungsdatum = dt.datetime(2021, 8, 1)
pfad = r'G:\Controlling\Python\Daten\Klarna\csv_files'
pfad_export = r'G:\Controlling\Python\Daten\Klarna\export'
pfad_eingelesen = r'G:\Controlling\Python\Daten\Klarna\eingelesen'

alte_dateien_loeschen()


# Erstellt eine Liste mit den im Pfad vorhandenen Dateien
dateien = os.listdir(pfad)

# Erstellt die csv-Datei "komplett.csv" mit sämtlichen Daten aus allen Dateien im Verzeichnis
for datei in dateien:
    if datei.lower().endswith('.csv'):
        with open(pfad + os.sep + datei, 'r') as ausgleich_daten:
            reader = csv.reader(ausgleich_daten, delimiter=";")
            kopfzeile = next(reader)
            daten.extend([zeile for zeile in reader])

        with open(pfad + os.sep + 'komplett.csv', 'w', newline='') as ergebnis_daten:
            writer = csv.writer(ergebnis_daten, delimiter=";")
            writer.writerow(kopfzeile)
            writer.writerows(daten)
        filename_von = os.path.join(pfad + os.sep, datei)
        filename_nach = os.path.join(pfad_eingelesen + os.sep, datei)
        shutil.move(filename_von, filename_nach)

komplett = (pfad + os.sep + 'komplett.csv')

df_sel = pd.read_csv(komplett,
                 encoding='cp1252',
                 sep=';',
                 usecols= [0, 2, 3, 11, 12, 23])

df_sel.columns = ['typ', 'verk_dat', 'erf_dat', 'betrag', 'whr', 'ref']

df_sel['verk_dat'] = pd.DatetimeIndex(df_sel['verk_dat'], dayfirst=True).date
df_sel['verk_dat'] = pd.to_datetime(df_sel['verk_dat'], format='%Y-%m-%d')
df_sel['verk_dat'] = df_sel['verk_dat'].apply(datum_vergleichen)
df_sel['erf_dat'] = pd.DatetimeIndex(df_sel['erf_dat'], dayfirst=True).date
df_sel['erf_dat'] = pd.to_datetime(df_sel['erf_dat'], format='%Y-%m-%d')
df_sel['erf_dat'] = df_sel['erf_dat'].apply(datum_vergleichen)
df_sel['ref'] = df_sel['ref'].astype(str)
df_sel['ref'] = 'Klarna ' + df_sel['ref']

df_day = df_sel.pivot_table(index=['ref', 'erf_dat'],
                            columns=['typ'],
                            values=['betrag'],
                            fill_value=0,
                            aggfunc='sum')

xw.view(df_day)
# Erstellen der CSV-Datei in CHF für FibuSync
df_day.to_csv(pfad_export + os.sep + "chf_Fibu_Sync.csv",
              header=False,
              # columns=['Datum', 'Fee', 'Retour', 'Verkauf',],
              index=True,
              sep=';')



alte_dateien_loeschen()