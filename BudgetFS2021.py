# Dieses Programm füllt die Umsätze nach Wargruppen 1- und 2stellig ab,
# diese wird für die Ermittlung der Einkaufsbudgets verwenden. Grundlage dafür ist die Umsatzliste 5.
# aus Apollon (Auswertung nach Verkaufspreisen)
#
# ACHTUNG: Dieses Jahr wurden zwei Perioden verwendet, darum 2 Dataframes, was in Zukunft nicht mehr der
# der Fall sein wird. Das zweite Dataframe df_2 muss in diesem Mall auskkommentiert werden, sonst werden
# die Daten dieses Dataframes dazugerechnet
# SG/30.7.2021

import pandas as pd
import xlwings as xw

COL = {0, 1, 2, 3, 4}  # Felder die aus dem Dataframe benötigt werden

wb = xw.Book(r'G:\controlling\python\Berichte\Budget\2107_Einkaufsbudget_FS2022.xlsm')

df = pd.read_csv(r'G:\controlling\python\daten\budget\budgetfs2022.csv', encoding='cp1252',
                 sep=';', index_col="FIL", usecols=COL, thousands="'")
df_2 = pd.read_csv(r'G:\controlling\python\daten\budget\budgetfs2022_2.csv', encoding='cp1252',
                   sep=';', index_col="FIL", usecols=COL, thousands="'")


df_alle = df.append(df_2)  # wenn zwei Perioden eingelesen werden sollen
# df_alle = df

df_alle['WG'] = df_alle['WAG'].apply(str)
df_alle['WG'] = df_alle['WG'].str.slice(0, 1)

df_wg1 = df_alle.loc[df_alle['WG'].isin(['1'])]
df_wg2 = df_alle.loc[df_alle['WG'].isin(['2'])]
df_wg3 = df_alle.loc[df_alle['WG'].isin(['3'])]
df_wg4 = df_alle.loc[df_alle['WG'].isin(['4'])]
df_wg5 = df_alle.loc[df_alle['WG'].isin(['5'])]
df_wg6 = df_alle.loc[df_alle['WG'].isin(['6'])]
df_wg9 = df_alle.loc[df_alle['WG'].isin(['9'])]

xw.sheets('WG1').range('a4').value = df_wg1.pivot_table(index=['FIL'], columns=['WAG'], values='UVKI', fill_value=0,
                                                        aggfunc='sum')
xw.sheets('WG1').range('o4').value = df_wg1.pivot_table(index=['FIL'], columns=['WAG'], values='UMEN', fill_value=0,
                                                        aggfunc='sum')
xw.sheets('WG2').range('a4').value = df_wg2.pivot_table(index=['FIL'], columns=['WAG'], values='UVKI', fill_value=0,
                                                        aggfunc='sum')
xw.sheets('WG2').range('o4').value = df_wg2.pivot_table(index=['FIL'], columns=['WAG'], values='UMEN', fill_value=0,
                                                        aggfunc='sum')
xw.sheets('WG3').range('a4').value = df_wg3.pivot_table(index=['FIL'], columns=['WAG'], values='UVKI', fill_value=0,
                                                        aggfunc='sum')
xw.sheets('WG3').range('o4').value = df_wg3.pivot_table(index=['FIL'], columns=['WAG'], values='UMEN', fill_value=0,
                                                        aggfunc='sum')
xw.sheets('WG4').range('a4').value = df_wg4.pivot_table(index=['FIL'], columns=['WAG'], values='UVKI', fill_value=0,
                                                        aggfunc='sum')
xw.sheets('WG4').range('o4').value = df_wg4.pivot_table(index=['FIL'], columns=['WAG'], values='UMEN', fill_value=0,
                                                        aggfunc='sum')
xw.sheets('WG5').range('a4').value = df_wg5.pivot_table(index=['FIL'], columns=['WAG'], values='UVKI', fill_value=0,
                                                        aggfunc='sum')
xw.sheets('WG5').range('o4').value = df_wg5.pivot_table(index=['FIL'], columns=['WAG'], values='UMEN', fill_value=0,
                                                        aggfunc='sum')
xw.sheets('WG6').range('a4').value = df_wg6.pivot_table(index=['FIL'], columns=['WAG'], values='UVKI', fill_value=0,
                                                        aggfunc='sum')
xw.sheets('WG6').range('o4').value = df_wg6.pivot_table(index=['FIL'], columns=['WAG'], values='UMEN', fill_value=0,
                                                        aggfunc='sum')
xw.sheets('WG9').range('a4').value = df_wg9.pivot_table(index=['FIL'], columns=['WAG'], values='UVKI', fill_value=0,
                                                        aggfunc='sum')
xw.sheets('WG9').range('o4').value = df_wg9.pivot_table(index=['FIL'], columns=['WAG'], values='UMEN', fill_value=0,
                                                        aggfunc='sum')
xw.sheets('Total').range('a4').value = df_alle.pivot_table(index=['FIL'], columns=['WG'], values='UVKI', fill_value=0,
                                                           aggfunc='sum')
xw.sheets('Total').range('o4').value = df_alle.pivot_table(index=['FIL'], columns=['WG'], values='UMEN', fill_value=0,
                                                           aggfunc='sum')

print(df_wg1)
print(df_alle.info())