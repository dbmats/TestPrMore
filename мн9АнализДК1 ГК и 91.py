import win32com.client as win32
import pandas as pd

# формирование Дт Кт оборотов

x = win32.GetObject(None, "Excel.Application")
wb = x.Workbooks("TestPrAn.xlsx")

d11 = pd.DataFrame(columns=['schD2', 'schK2', 'god', 'mes', 'sum']) # создание пустой заготовки
d22 = pd.DataFrame(columns=['schD2', 'schK2', 'cD1', 'cK1', 'god', 'mes', 'sum'])
for i in ['PrTest1.csv',
          'PrTest2.csv'
          ]:
    dp = pd.read_csv(i)
    dp = dp.fillna(0)
    dp1 = dp[['schD2', 'schK2', 'god', 'mes', 'sum']]
    d11 = d11.append(dp1)
    dp91 = dp[['schD2', 'schK2', 'cD1', 'cK1', 'god', 'mes', 'sum']]
    dp91d = dp91.loc[dp91['schD2'] == 91]
    dp91d = dp91d.loc[~dp91d['schK2'].isin([91, 99])]
    dp91k = dp91.loc[dp91['schK2'] == 91]
    dp91k = dp91k.loc[~dp91k['schD2'].isin([91, 99])]
    dp91k = dp91k.append(dp91d)
    d22 = d22.append(dp91k)
dp1 = d11
dp911 = d22

# ГК
tb = pd.pivot_table(dp1, values=['sum'],
                    index=['schD2', 'schK2', 'god', 'mes'],
                    aggfunc={'sum': sum})
tb.reset_index(inplace=True)
tb['sumK'] = 0.0
tb.columns = ['sch', 'ksch', 'god', 'mes', 'sumD', 'sumK']

tb1 = pd.pivot_table(dp1, values=['sum'],
                     index=['schK2', 'schD2', 'god', 'mes'],
                     aggfunc={'sum': sum})
tb1.reset_index(inplace=True)
tb1['sumD'] = 0.0
tb1.columns = ['sch', 'ksch', 'god', 'mes', 'sumK', 'sumD']
tb1 = tb1[['sch', 'ksch', 'god', 'mes', 'sumD', 'sumK']]

fr = [tb, tb1]
pp = pd.concat(fr)

pp1 = pd.pivot_table(pp, values=['sumD', 'sumK'],
                     index=['sch', 'ksch'],
                     columns=['god', 'mes'],
                     aggfunc={'sumD': sum, 'sumK': sum},
                     margins= True)
pp1.reset_index(inplace=True)
pp1 = pp1.fillna(0)
ws = wb.Worksheets("ГКдк")
StartRow = 6
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(pp1.index) - 1,
                  StartCol + len(pp1.columns) - 1)).Value = pp1.values

# сч.91
dp91 = dp911.loc[dp911['schD2'] == 91]
dp91 = dp91.loc[~dp91['schK2'].isin([91, 99])]
tb = pd.pivot_table(dp91, values=['sum'],
                    index=['cD1','schK2', 'god'],
                    aggfunc={'sum': sum})
tb.reset_index(inplace=True)
tb['sumK'] = 0.0
tb.columns = ['cubk', 'Ksch', 'god', 'sumD', 'sumK']

dp91 = dp911.loc[dp911['schK2'] == 91]
dp91 = dp91.loc[~dp91['schD2'].isin([91, 99])]
tb1 = pd.pivot_table(dp91, values=['sum'],
                     index=['cK1', 'schD2', 'god'],
                     aggfunc={'sum': sum})
tb1.reset_index(inplace=True)
tb1['sumD'] = 0.0
tb1.columns = ['cubk', 'Ksch', 'god', 'sumK', 'sumD']
tb1 = tb1[['cubk', 'Ksch', 'god', 'sumD', 'sumK']]

fr = [tb, tb1]
pp = pd.concat(fr)
pp.reset_index(inplace=True)

pp3 = pd.pivot_table(pp, values=['sumD', 'sumK'],
                     index=['cubk'],
                     columns=['god', 'Ksch'],
                     aggfunc={'sumD': sum, 'sumK': sum},
                     margins=True)
pp3.reset_index(inplace=True)
pp3 = pp3.fillna(0)

ws = wb.Worksheets("сч.91")
StartRow = 6
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(pp3.index) - 1,
                  StartCol + len(pp3.columns) - 1)).Value = pp3.values
pp3.to_excel("output.xlsx")

# print(pp3)
# print(dp.head(5))
# print(dp.columns)
# print(dp.shape)
# print(tb.dtypes)
