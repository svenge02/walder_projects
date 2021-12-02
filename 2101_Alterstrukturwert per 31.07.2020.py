import pandas as pd
import xlwings as xw


wb = xw.Book(r'G:\controlling\python\Berichte\Altersstruktur\2021_Altersstruktur.xlsx')

df = pd.read_csv(r'G:\controlling\python\daten\Altersstruktur\Altersstruktur2021.csv', encoding='cp1252',
                 sep=';', thousands="'")

filialliste = ['098', '099', '101', '102', '103', '105', '107', '114', '117', '120',
               '126', '127', '141', '210', '215', '216', '221', '224', '228',
               '233', '234', '235', '240', '244', '246', '342', '406', '408', '411', '780',
               '551', '552', '553', '573', '574', '782', '790', '793']

warengruppen = [0, 1, 2, 3, 4, 5, 6, 9]

spalten = ['BEKN', 'asw']


def abschlag(row):
    # Definierte Abschlagssätze nach Warengruppen
    wg0_stand = 0.1
    wg0_nw = 0.0
    wg0_05j = 0.0
    wg0_10j = 0.0
    wg0_15j = 0.0
    wg0_20j = 0.0
    wg0_alt = 0.0
    wg1_stand = 0.0
    wg1_nw = 0.0
    wg1_05j = 0.1
    wg1_10j = 0.2
    wg1_15j = 0.3
    wg1_20j = 0.5
    wg1_alt = 0.9
    wg2_stand = 0.0
    wg2_nw = 0.0
    wg2_05j = 0.15
    wg2_10j = 0.25
    wg2_15j = 0.40
    wg2_20j = 0.60
    wg2_alt = 0.9

    wg3_stand = 0.0
    wg3_nw = 0.0
    wg3_05j = 0.1
    wg3_10j = 0.2
    wg3_15j = 0.3
    wg3_20j = 0.5
    wg3_alt = 0.9

    wg4_stand = 0.0
    wg4_nw = 0.0
    wg4_05j = 0.15
    wg4_10j = 0.25
    wg4_15j = 0.4
    wg4_20j = 0.6
    wg4_alt = 0.9

    wg5_stand = 0.0
    wg5_nw = 0.0
    wg5_05j = 0.1
    wg5_10j = 0.2
    wg5_15j = 0.3
    wg5_20j = 0.5
    wg5_alt = 0.9

    wg6_stand = 0.0
    wg6_nw = 0.0
    wg6_05j = 0.1
    wg6_10j = 0.2
    wg6_15j = 0.3
    wg6_20j = 0.5
    wg6_alt = 0.9

    wg9_stand = 0.0
    wg9_nw = 0.0
    wg9_05j = 0.15
    wg9_10j = 0.25
    wg9_15j = 0.4
    wg9_20j = 0.6
    wg9_alt = 0.9

    # Warengruppe 0
    if row.WAG == 0 and row.Kennung == 'stand':
        return wg0_stand
    elif row.WAG == 0 and row.Kennung == 'nw':
        return wg0_nw
    elif row.WAG == 0 and row.Kennung == 'A05j':
        return wg0_05j
    elif row.WAG == 0 and row.Kennung == 'A10j':
        return wg0_10j
    elif row.WAG == 0 and row.Kennung == 'A15j':
        return wg0_15j
    elif row.WAG == 0 and row.Kennung == 'A20j':
        return wg0_20j
    elif row.WAG == 0 and row.Kennung == 'alt':
        return wg0_alt

    # Warengruppe 1
    if row.WAG == 1 and row.Kennung == 'stand':
        return wg1_stand
    elif row.WAG == 1 and row.Kennung == 'nw':
        return wg1_nw
    elif row.WAG == 1 and row.Kennung == 'A05j':
        return wg1_05j
    elif row.WAG == 1 and row.Kennung == 'A10j':
        return wg1_10j
    elif row.WAG == 1 and row.Kennung == 'A15j':
        return wg1_15j
    elif row.WAG == 1 and row.Kennung == 'A20j':
        return wg1_20j
    elif row.WAG == 1 and row.Kennung == 'alt':
        return wg1_alt

    # Warengruppe 2
    if row.WAG == 2 and row.Kennung == 'stand':
        return wg2_stand
    elif row.WAG == 2 and row.Kennung == 'nw':
        return wg2_nw
    elif row.WAG == 2 and row.Kennung == 'A05j':
        return wg2_05j
    elif row.WAG == 2 and row.Kennung == 'A10j':
        return wg2_10j
    elif row.WAG == 2 and row.Kennung == 'A15j':
        return wg2_15j
    elif row.WAG == 2 and row.Kennung == 'A20j':
        return wg2_20j
    elif row.WAG == 2 and row.Kennung == 'alt':
        return wg2_alt

    # Warengruppe 3
    if row.WAG == 3 and row.Kennung == 'stand':
        return wg3_stand
    elif row.WAG == 3 and row.Kennung == 'nw':
        return wg3_nw
    elif row.WAG == 3 and row.Kennung == 'A05j':
        return wg3_05j
    elif row.WAG == 3 and row.Kennung == 'A10j':
        return wg3_10j
    elif row.WAG == 3 and row.Kennung == 'A15j':
        return wg3_15j
    elif row.WAG == 3 and row.Kennung == 'A20j':
        return wg3_20j
    elif row.WAG == 3 and row.Kennung == 'alt':
        return wg3_alt

    # Warengruppe 4
    if row.WAG == 4 and row.Kennung == 'stand':
        return wg4_stand
    elif row.WAG == 4 and row.Kennung == 'nw':
        return wg4_nw
    elif row.WAG == 4 and row.Kennung == 'A05j':
        return wg4_05j
    elif row.WAG == 4 and row.Kennung == 'A10j':
        return wg4_10j
    elif row.WAG == 4 and row.Kennung == 'A15j':
        return wg4_15j
    elif row.WAG == 4 and row.Kennung == 'A20j':
        return wg4_20j
    elif row.WAG == 4 and row.Kennung == 'alt':
        return wg4_alt

    # Warengruppe 5
    if row.WAG == 5 and row.Kennung == 'stand':
        return wg5_stand
    elif row.WAG == 5 and row.Kennung == 'nw':
        return wg5_nw
    elif row.WAG == 5 and row.Kennung == 'A05j':
        return wg5_05j
    elif row.WAG == 5 and row.Kennung == 'A10j':
        return wg5_10j
    elif row.WAG == 5 and row.Kennung == 'A15j':
        return wg5_15j
    elif row.WAG == 5 and row.Kennung == 'A20j':
        return wg5_20j
    elif row.WAG == 5 and row.Kennung == 'alt':
        return wg5_alt

    # Warengruppe 6
    if row.WAG == 6 and row.Kennung == 'stand':
        return wg6_stand
    elif row.WAG == 6 and row.Kennung == 'nw':
        return wg6_nw
    elif row.WAG == 6 and row.Kennung == 'A05j':
        return wg6_05j
    elif row.WAG == 6 and row.Kennung == 'A10j':
        return wg6_10j
    elif row.WAG == 6 and row.Kennung == 'A15j':
        return wg6_15j
    elif row.WAG == 6 and row.Kennung == 'A20j':
        return wg6_20j
    elif row.WAG == 6 and row.Kennung == 'alt':
        return wg6_alt

    # Warengruppe 9
    if row.WAG == 9 and row.Kennung == 'stand':
        return wg9_stand
    elif row.WAG == 9 and row.Kennung == 'nw':
        return wg9_nw
    elif row.WAG == 9 and row.Kennung == 'A05j':
        return wg9_05j
    elif row.WAG == 9 and row.Kennung == 'A10j':
        return wg9_10j
    elif row.WAG == 9 and row.Kennung == 'A15j':
        return wg9_15j
    elif row.WAG == 9 and row.Kennung == 'A20j':
        return wg9_20j
    elif row.WAG == 9 and row.Kennung == 'alt':
        return wg9_alt
    else:

        return 0.0


print(df.shape)

# df['UVKI'] = pd.to_numeric(df['UVKI'])
df_neu = df
# df.drop(df.WAGBEZ['Reparaturen'])

Saison_Zuteilung = {'09': 'stand',
                    '58': 'nw',
                    '57': 'nw',
                    '56': 'nw',
                    '55': 'nw',
                    '54': 'A05j',
                    '53': 'A10j',
                    '52': 'A15j',
                    '51': 'A20j',
                    '50': 'alt',
                    '49': 'alt',
                    '48': 'alt',
                    '47': 'alt',
                    '46': 'alt',
                    '45': 'alt',
                    '44': 'alt',
                    '41': 'alt',
                    '42': 'alt',
                    '43': 'alt',
                    '40': 'alt',
                    '39': 'alt',
                    '38': 'alt',
                    '36': 'alt',
                    '35': 'alt',
                    '34': 'alt',
                    '33': 'alt',
                    '32': 'alt',
                    '31': 'alt',
                    '30': 'alt',
                    '29': 'alt',
                    '28': 'alt',
                    '27': 'alt',
                    '26': 'alt',
                    '25': 'alt',
                    '24': 'alt',
                    '23': 'alt',
                    '22': 'alt',
                    '21': 'alt',
                    '20': 'alt',
                    '17': 'alt',
                    '08': 'alt'}

print(df_neu)

df_neu['SAISON'] = df_neu['SAISON'].str.split(" - ", n=1, expand=True)
df_neu['FilialeNr'] = df_neu['FILIALE'].str.replace("Filiale ", "")
# df_neu['Kennung'] = df_neu.where(df_neu['SAISON'] == '09', 'stand')
df_neu['Kennung'] = [Saison_Zuteilung[i] for i in df['SAISON']]
df_neu['Abschlag'] = 0

df.loc[:, 'Abschlag'] = df_neu.apply(abschlag, axis=1)
df_neu['asw'] = df_neu['BEKN'] * (1- df_neu['Abschlag'])

# df_select = (df_neu['FilialeNr'] == '095')
# df_select =df_neu.loc[df_neu['FilialeNr'].isin(['101', '102'])] #richtiger Suchstring
df_select = df_neu.loc[df_neu['FilialeNr'].isin(filialliste)]
df_un_select = df_neu.loc[~df_neu['FilialeNr'].isin(filialliste)]


df_wg0 = df_select.loc[df_select['WAG'].isin([0])]
df_wg1 = df_select.loc[df_select['WAG'].isin([1])]
df_wg2 = df_select.loc[df_select['WAG'].isin([2])]
df_wg3 = df_select.loc[df_select['WAG'].isin([3])]
df_wg4 = df_select.loc[df_select['WAG'].isin([4])]
df_wg5 = df_select.loc[df_select['WAG'].isin([5])]
df_wg6 = df_select.loc[df_select['WAG'].isin([6])]
df_wg9 = df_select.loc[df_select['WAG'].isin([9])]




print(df_neu.info())


# xw.view(df_neu.groupby(['WAG', 'Kennung'])['BEKN'].agg('sum'))



# xw.view(df_select.groupby(['Kennung'])['BMEN','BEKN', 'asw'].agg('sum'))
# xw.view(df_select.groupby(['FILIALE'])['BMEN','BEKN', 'asw'].agg('sum'))
#xw.view(df_un_select.groupby(['FILIALE'])['BEKN', 'asw'].agg('sum'))

xw.sheets['aktuell'].range('c5').value = df_select.groupby(['Kennung'])['BMEN','BVKS', 'BVKA','BEKN', 'asw'].agg('sum')

xw.sheets['aktuell'].range('B16').value = df_wg0.groupby(['WAG', 'Kennung'])['BMEN','BVKS', 'BVKA','BEKN', 'asw'].agg('sum')
xw.sheets['aktuell'].range('B27').value = df_wg1.groupby(['WAG', 'Kennung'])['BMEN','BVKS', 'BVKA','BEKN', 'asw'].agg('sum')
xw.sheets['aktuell'].range('B38').value = df_wg2.groupby(['WAG', 'Kennung'])['BMEN','BVKS', 'BVKA','BEKN', 'asw'].agg('sum')
xw.sheets['aktuell'].range('B49').value = df_wg3.groupby(['WAG', 'Kennung'])['BMEN','BVKS', 'BVKA','BEKN', 'asw'].agg('sum')
xw.sheets['aktuell'].range('B60').value = df_wg4.groupby(['WAG', 'Kennung'])['BMEN','BVKS', 'BVKA','BEKN', 'asw'].agg('sum')
xw.sheets['aktuell'].range('B71').value = df_wg5.groupby(['WAG', 'Kennung'])['BMEN','BVKS', 'BVKA','BEKN', 'asw'].agg('sum')
xw.sheets['aktuell'].range('B82').value = df_wg6.groupby(['WAG', 'Kennung'])['BMEN','BVKS', 'BVKA','BEKN', 'asw'].agg('sum')
xw.sheets['aktuell'].range('B93').value = df_wg9.groupby(['WAG', 'Kennung'])['BMEN','BVKS', 'BVKA','BEKN', 'asw'].agg('sum')

xw.sheets['aktuell'].range('C105').value = df_select.groupby(['FILIALE'])['BMEN','BVKS', 'BVKA','BEKN', 'asw'].agg('sum')
xw.sheets['aktuell'].range('C149').value = df_un_select.groupby(['FILIALE'])['BMEN','BVKS', 'BVKA','BEKN', 'asw'].agg('sum')