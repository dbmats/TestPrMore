import numpy as np
import pandas as pd
from win32com.client import GetObject

# Аналитические процедуры

dd11 = pd.DataFrame(columns=['schK2', 'cK1', 'sum', 'god', 'mes']) # создание пустой заготовки
dd22 = pd.DataFrame(columns=['schD2', 'schK2', 'cD1', 'cD2', 'cD3', 'cK1', 'cK2' , 'god', 'mes', 'sum'])
dd33 = pd.DataFrame(columns=['cD1', 'cD2', 'cK1', 'god', 'mes', 'sum'])
for i in ['PrTest1.csv',
          'PrTest2.csv'
          ]:
    dp = pd.read_csv(i)
    dp = dp.fillna(0)
    dp1 = dp[['schK2', 'cK1', 'sum', 'god', 'mes']]
    dp1 = dp.loc[dp['schK2'] == 2]
    dd11 = dd11.append(dp1)
    dp2 = dp.loc[dp['schK2'].isin([60, 76])]
    dp2 = dp2.loc[dp2['schD2'].isin([8, 7, 10, 15, 20, 23, 25, 26, 29, 41, 44, 91, 97, '08', '07'])]
    dd22 = dd22.append(dp2)
    dp3 = dp.loc[dp['schD2'].isin([62])]
    dp3 = dp3.loc[dp3['schK2'].isin([90])]
    dd33 = dd33.append(dp3)
dp1 = dd11
dp3 = dd22
dp4 = dd33


# # начисление амортизации
tb2 = pd.pivot_table(dp1, values=['sum'],
                    index=['cK1'],
                    aggfunc={'sum': sum},
                    columns=['god', 'mes'], margins=True)
tb2.reset_index(inplace=True)
tb2 = tb2.fillna(0)

# закупки К60,76
tb4 = pd.pivot_table(dp3, values=['sum'],
                     index=['schD2', 'schK2', 'cD1', 'cD2', 'cD3', 'cK1', 'cK2' , 'god', 'mes'],
                     aggfunc={'sum': sum},)
tb4.reset_index(inplace=True)
tb4 = tb4.fillna(0)
tb4.to_csv('закупки.csv', index=False)

# # доходы Д62К90
tb5 = pd.pivot_table(dp4, values=['sum'],
                     index=['cD1', 'cD2', 'cK1', 'god', 'mes'],
                     aggfunc={'sum': sum},)
tb5.reset_index(inplace=True)
tb5 = tb5.fillna(0)
tb5.to_csv('выручка.csv', index=False)

# Выгрузка сделанных анализов в РД
x = GetObject(None, "Excel.Application")
wb = x.Workbooks("TestPrAn.xlsx")

ws = wb.Worksheets("аморт")
StartRow = 6
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(tb2.index) - 1,
                  StartCol + len(tb2.columns) - 1)).Value = tb2.values

ws = wb.Worksheets("закуп")
StartRow = 2
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(tb4.index) - 1,
                  StartCol + len(tb4.columns) - 1)).Value = tb4.values

ws = wb.Worksheets("выр")
StartRow = 2
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(tb5.index) - 1,
                  StartCol + len(tb5.columns) - 1)).Value = tb5.values

# tb3.to_excel("output.xlsx")

# print(tb4)
# print(dp.head(5))
# print(dp.columns)
# print(tb.shape)
# print(dp.dtypes)
