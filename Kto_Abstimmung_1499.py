"""
Kto_Abstimmung_1499.py
This program is supposed to reconcile one special account which is used for the payments from the shops. Usually
the same amount appears on both sides. In that case the booking should be deleted so that only the ones with equal
amount remain open
sg/15.06.21
"""

import pandas as pd
import xlwings as xw


reihen_zaehler_1 = 0
reihen_zaehler_2 = 0
gesucht_betrag = 0.0

columns_not_used = [1, 7, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35,
                    36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57,
                    58, 59, 60, 61, 62, 63, 64]

col = ['abac_nr', 'booking_date', 'account_debit', 'accoutnt_credit', 'booking_text', 'booking_amount',
       'booking_type', 'shop_debit', 'shop_credit', 'bookinb_nr', 'e_nr', 'currency_1', 'currency_2']

df = pd.read_csv(r'G:\controlling\python\daten\Abacus\Kontoabstimmungen\Kto1499.csv', encoding='cp1252',
                 sep=',')

df.drop(df.columns[tuple([columns_not_used])],axis=1,inplace=True)
df.columns = col


print(df.shape[0])


""""
for reihen_zaehler_1 in range(df.shape[0]+1):
    gesucht_betrag = df['kto_betrag'][reihen_zaehler_1]
    for reihen_zaehler_2 in range(10):
        reihen_zaehler_2 += 1
        if df['abgestimmt'][reihen_zaehler_2] != "X":
            if gesucht_betrag == df['kto_betrag'][reihen_zaehler_2]:
                df['abgestimmt'][reihen_zaehler_1] = "X"
                df['abgestimmt'][reihen_zaehler_2] = "X"
                break
            else:
                reihen_zaehler_2 += 1
        else:
            reihen_zaehler_2 += 1

"""
xw.view(df)
