import win32com.client as win32
import pandas as pd

x = win32.GetObject(None, "Excel.Application")
wb = x.Workbooks("TestPrAn.xlsx")

d11 = pd.DataFrame(columns=['cubk', 'sch', 'Ksch', 'god', 'mes', 'sumD', 'sumK'])
for i in ['PrTest1.csv',
          'PrTest2.csv'
          ]:
    dp = pd.read_csv(i)
    dp = dp.fillna(0)

    dp90 = dp.loc[dp['schD2'] == 90]
    dp90 = dp90.loc[~dp90['schK2'].isin([90, 99])]
    tb = pd.pivot_table(dp90, values=['sum'],
                    index=['cD1', 'schD', 'schK2', 'god', 'mes'],
                    aggfunc={'sum': sum})
    tb.reset_index(inplace=True)
    tb['sumK'] = 0.0
    tb.columns = ['cubk', 'sch', 'Ksch', 'god', 'mes', 'sumD', 'sumK']

    dp90 = dp.loc[dp['schK2'] == 90]
    dp90 = dp90.loc[~dp90['schD2'].isin([90, 99])]
    tb1 = pd.pivot_table(dp90, values=['sum'],
                     index=['cK1', 'schK', 'schD2', 'god', 'mes'],
                     aggfunc={'sum': sum})
    tb1.reset_index(inplace=True)
    tb1['sumD'] = 0.0
    tb1.columns = ['cubk', 'sch', 'Ksch', 'god', 'mes', 'sumK', 'sumD']
    tb1 = tb1[['cubk', 'sch', 'Ksch', 'god', 'mes', 'sumD', 'sumK']]

    fr = [tb, tb1]
    pp = pd.concat(fr)
    pp.reset_index(inplace=False)

    d11 = d11.append(pp)
dp = d11

dp.to_csv('90.csv', index=False)

ws = wb.Worksheets("сч.90")
StartRow = 2
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(dp.index) - 1,
                  StartCol + len(dp.columns) - 1)).Value = dp.values


# print(dp90)
print(dp.head(5))
# print(dp.columns)
# print(dp.shape)
# print(dp.dtypes)
