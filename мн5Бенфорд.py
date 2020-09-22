import pandas as pd
from win32com.client import GetObject

x = GetObject(None, "Excel.Application")
wb = x.Workbooks("TestPr.xlsx")

pp1 = pd.DataFrame(columns=['x1', 'mes', 'data'])
pp2 = pd.DataFrame(columns=['x2', 'mes', 'data'])
pp3 = pd.DataFrame(columns=['xx', 'mes', 'data'])
for i in ['PrTest2.csv',
          'PrTest1.csv'
          ]:
    dp = pd.read_csv(i)
    dp = dp.fillna(0)

    # dp = dp.loc[dp['god'] == 2019]
    # dp = dp.loc[~dp['x1'].isin([0])]

    tb1 = pd.pivot_table(dp, values=['data'],
                         index=['x1', 'mes'],
                         aggfunc={'data': 'count'})
    tb1.reset_index(inplace=True)
    tb1 = tb1.fillna(0)
    pp1 = pp1.append(tb1)

    tb2 = pd.pivot_table(dp, values=['data'],
                         index=['x2', 'mes'],
                         aggfunc={'data': 'count'})
    tb2.reset_index(inplace=True)
    tb2 = tb2.fillna(0)
    pp2 = pp2.append(tb2)

    tb3 = pd.pivot_table(dp, values=['data'],
                         index=['xx', 'mes'],
                         aggfunc={'data': 'count'})
    tb3.reset_index(inplace=True)
    tb3 = tb3.fillna(0)
    pp3 = pp3.append(tb3)


tb1 = pd.pivot_table(pp1, values=['data'],
                     index=['x1'],
                     columns=['mes'],
                     aggfunc={'data': sum}, margins=True)
tb1.reset_index(inplace=True)
tb1 = tb1.fillna(0)
ws = wb.Worksheets("Бенфорд")
StartRow = 6
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(tb1.index) - 1,
                  StartCol + len(tb1.columns) - 1)).Value = tb1.values

tb2=pd.pivot_table(pp2, values=['data'],
                   index=['x2'],
                   columns=['mes'],
                   aggfunc={'data': sum}, margins=True)
tb2.reset_index(inplace=True)
tb2 = tb2.fillna(0)
ws = wb.Worksheets("Бенфорд")
StartRow = 19
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(tb2.index) - 1,
                  StartCol + len(tb2.columns) - 1)).Value = tb2.values

tb = pd.pivot_table(pp3, values=['data'],
                    index=['xx'],
                    columns=['mes'],
                    aggfunc={'data': sum}, margins=True)
tb.reset_index(inplace=True)
tb = tb.fillna(0)
ws = wb.Worksheets("Бенфорд")
StartRow = 33
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(tb.index) - 1,
                  StartCol + len(tb.columns) - 1)).Value = tb.values

# dp = dp.loc[dp['mes'] == 3]
# dp = dp.loc[dp['x2'] == 0]
# tb2.to_excel("output.xlsx")
# print(dp[['sum', 'xxx', 'xx', 'x1', 'x2']])
# print(tb1)
# print(dp.dtypes)
