"""
Dieses Skript wertet die Klarna-Daten aus.
Die Übersicht zeigt auf, wieviele Käufe über Klarna getätigt wurden,
was zurückgeschickt wurde und wie hoch die Gebühren waren.
SG/21.06.21
"""

import pandas as pd
import xlwings as xw

current_month = 10
current_year = 2021

COL = ['type', 'sale_date', 'sale_month',  'sale_year', 'capture_date', 'amount', 'payout_date',
       'payout_month', 'payout_year', 'payout_amount']


wb = xw.Book(r'G:\controlling\python\Berichte\2021_Klarna_Übersicht.xlsx')

df = pd.read_csv(r'G:\controlling\python\daten\Klarna\reports\klarnaabrechnung.csv', sep=';')
df_sel = pd.DataFrame(columns=COL)

df_sel['type'] = df['type']
df_sel['sale_date'] = pd.DatetimeIndex(df['sale_date'], dayfirst=True)
df_sel['sale_month'] = pd.DatetimeIndex(df_sel['sale_date']).month
df_sel['sale_year'] = pd.DatetimeIndex(df_sel['sale_date']).year
df_sel['capture_date'] = pd.DatetimeIndex(df['capture_date'], dayfirst=True)
df_sel['capture_month'] = pd.DatetimeIndex(df_sel['capture_date']).month
df_sel['caputre_year'] = pd.DatetimeIndex(df_sel['capture_date']).year
df_sel['amount'] = df['amount']
df_sel['payout_date'] = pd.DatetimeIndex(df['payout_date'], dayfirst=True)
df_sel['payout_month'] = pd.DatetimeIndex(df_sel['payout_date']).month
df_sel['payout_year'] = pd.DatetimeIndex(df_sel['payout_date']).year
df_sel['payout_amount'] = df_sel['amount']
# df_sel.loc[:,'payout_amount'] = df_sel.apply(set_payout_amount, axis=1)
# df_sel_recent = df_sel.loc[(df_sel.payout_month <= actual_month)]


empty_payout_date = df_sel['payout_date'].isna()
df_sel.loc[empty_payout_date, 'payout_amount'] = 0.0


present_year = df_sel['payout_date'].dt.year <= current_year
future_month = df_sel['payout_date'].dt.month > current_month
df_sel.loc[present_year & future_month, 'payout_amount'] = 0.0

xw.sheets('2021').range('B6').value = df_sel.pivot_table(index=['sale_month'], columns=['type'],values=['amount'],
                                                         fill_value=0, aggfunc='sum')

xw.sheets('2021').range('B26').value = df_sel.pivot_table(index=['capture_month'], columns=['type'],
                                                          values=['amount'], fill_value=0, aggfunc='sum')




print(df_sel)

