import win32com.client as win32
import pandas as pd

# формирование Дт Кт оборотов

x = win32.GetObject(None, "Excel.Application")
wb = x.Workbooks("TestPrAn.xlsx")

d11 = pd.DataFrame(columns=['cub1', 'cub2', 'sch', 'Ksch', 'god', 'mes', 'sumD', 'sumK'])
for i in ['PrTest1.csv',
          'PrTest2.csv'
          ]:
    dp = pd.read_csv(i)
    dp = dp.fillna(0)

    dpDK = dp.loc[dp['schD2'].isin([46, 58, 59, 60, 62, 63, 66, 67, 71, 73, 75, 76])]
    tb = pd.pivot_table(dpDK, values=['sum'],
                        index=['cD1', 'cD2',
                               'schD2', 'schK2', 'god', 'mes'],
                        aggfunc={'sum': sum})
    tb.reset_index(inplace=True)
    tb['sumK'] = 0.0
    tb.columns = ['cub1', 'cub2', 'sch', 'Ksch', 'god', 'mes', 'sumD', 'sumK']

    dpDK = dp.loc[dp['schK2'].isin([46, 58, 59, 60, 62, 63, 66, 67, 71, 73, 75, 76])]
    tb1 = pd.pivot_table(dpDK, values=['sum'],
                     index=['cK1', 'cK2',
                            'schK2', 'schD2', 'god', 'mes'],
                     aggfunc={'sum': sum})
    tb1.reset_index(inplace=True)
    tb1['sumD'] = 0.0
    tb1.columns = ['cub1', 'cub2', 'sch', 'Ksch', 'god', 'mes', 'sumK', 'sumD']
    tb1 = tb1[['cub1', 'cub2', 'sch', 'Ksch', 'god', 'mes', 'sumD', 'sumK']]

    fr = [tb, tb1]
    pp = pd.concat(fr)
    pp.reset_index(inplace=False)
    d11 = d11.append(pp)
dp = d11

dp.to_csv('60.csv', index=False)
ws = wb.Worksheets("ДзКз")
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
