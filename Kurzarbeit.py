# für Dataframes
import pandas as pd
# für die Integration in Excel
import xlwings as xw


# Definition der Konstanten

MAN = "L001"
LOHN_DATUM = "30.11.2021"
NR1 = 1


def loa_ermitteln(row):
  if row.besch == 0:
      return 1550
  else:
      return 1500

def betrag_ermitteln(row):
    if row.besch == 0:
        return row.loa1550
    else:
        return row.loa1500


columns = []

# wb = xw.Book(r'G:\controlling\python\Berichte\Wochenumsatz\Wochenumsatz2021_1.xls')
df = pd.read_csv(r'G:\controlling\python\daten\lohn\kua.csv', encoding='cp1252',
                 sep=';', thousands="'")


col = ['mandant', 'pnr', 'datum', 'nr1', 'loa', 'dm_1', 'dm_2', 'dm_3', 'betrag', 'dm_4', 'dm_5', 'dm_6', 'dm_7'
           'dm_8', 'dm_9', 'dm_10', 'dm_11', 'dm_12', 'mandant_2']

df_neu = pd.DataFrame(columns=col)
df_neu2 = pd.DataFrame(columns=col)
df_neu3 = pd.DataFrame(columns=col)


df_neu['pnr'] = df['pnr']
df_neu['datum'] = LOHN_DATUM
df_neu.loc[:, 'loa'] = df.apply(loa_ermitteln, axis=1)
df_neu.loc[:, 'betrag'] = df.apply(betrag_ermitteln, axis=1)
df_neu['mandant'] = MAN
df_neu['mandant_2'] = MAN
df_neu['nr1'] = NR1

print(df)
print(df.info())
print(df_neu)
print(df_neu.info())


df_neu.drop(df_neu[df_neu.betrag == 0.0].index, inplace=True)

df_neu.to_csv(r'G:\controlling\python\daten\lohn\kua_neu.csv', header = False, sep=',', index = False)

df_neu2['pnr'] = df['pnr']
df_neu2['datum'] = LOHN_DATUM
df_neu2['loa'] = 1200
df_neu2['betrag'] = df['loa1200']
df_neu2['mandant'] = MAN
df_neu2['mandant_2'] = MAN
df_neu2['nr1'] = NR1

df_neu2.drop(df_neu2[df_neu2.betrag == 0.0].index, inplace=True)

df_neu2.to_csv(r'G:\controlling\python\daten\lohn\kua_neu.csv', mode = 'a',  header = False, sep=',', index = False)

df_neu3['pnr'] = df['pnr']
df_neu3['datum'] = LOHN_DATUM
df_neu3['loa'] = 1205
df_neu3['betrag'] = df['loa1205']
df_neu3['mandant'] = MAN
df_neu3['mandant_2'] = MAN
df_neu3['nr1'] = NR1


df_neu3.drop(df_neu3[df_neu3.betrag == 0.0].index, inplace=True)

df_neu3.to_csv(r'G:\controlling\python\daten\lohn\kua_neu.csv', mode = 'a',  header = False, sep=',', index = False)