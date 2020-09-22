import win32com.client as win32
import pandas as pd

x = win32.GetObject(None, "Excel.Application")
wb = x.Workbooks("TestPrAn.xlsx")

d11 = pd.DataFrame(columns=['cub1', 'cub2', 'cub3', 'sch', 'Ksch', 'god', 'mes', 'sumD', 'sumK'])
for i in ['PrTest1.csv',
          'PrTest2.csv'
          ]:
    dp = pd.read_csv(i)
    dp = dp.fillna(0)

    dp20 = dp.loc[dp['schD2'].isin([20,
                                    21, 23, 25, 26, 28, 29, 43, 44, 8, '08'
                                    ])]
    tb = pd.pivot_table(dp20, values=['sum'],
                        index=['cD1', 'cD2', 'cD3',
                               'schD2', 'schK2', 'god', 'mes'],
                        aggfunc={'sum': sum})
    tb.reset_index(inplace=True)
    tb['sumK'] = 0.0
    tb.columns = ['cub1', 'cub2', 'cub3', 'sch', 'Ksch', 'god', 'mes', 'sumD', 'sumK']

    dp90 = dp.loc[dp['schK2'].isin([20,
                                    21, 23, 25, 26, 28, 29, 43, 44, 8, '08'
                                    ])]
    tb1 = pd.pivot_table(dp90, values=['sum'],
                         index=['cK1', 'cK2', 'cK3',
                                'schK2', 'schD2', 'god', 'mes'],
                         aggfunc={'sum': sum})
    tb1.reset_index(inplace=True)
    tb1['sumD'] = 0.0
    tb1.columns = ['cub1', 'cub2', 'cub3', 'sch', 'Ksch', 'god', 'mes', 'sumK', 'sumD']
    tb1 = tb1[['cub1', 'cub2', 'cub3', 'sch', 'Ksch', 'god', 'mes', 'sumD', 'sumK']]

    fr = [tb, tb1]
    pp = pd.concat(fr)
    pp.reset_index(inplace=False)
    d11 = d11.append(pp)
dp = d11

dp.to_csv('20.csv', index=False)

ws = wb.Worksheets("08,20-44")
StartRow = 2
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(dp.index) - 1,
                  StartCol + len(dp.columns) - 1)).Value = dp.values


# print(pp3)
print(dp.head(5))
# print(dp.columns)
# print(dp.shape)
# print(tb.dtypes)
