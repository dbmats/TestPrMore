from win32com.client import GetObject
import pandas as pd
import pickle

muk = 41647
dp = pd.read_csv('PrTest1.csv')
dp['muk'] = dp['sum'].abs() > muk
dp.to_csv('PrTest1.csv', index=False)

dp = pd.read_csv('PrTest2.csv')
dp['muk'] = dp['sum'].abs() > muk
dp.to_csv('PrTest2.csv', index=False)

# формирование ГК
pp = pd.DataFrame(columns=['schD2','schK2', 'muk', 'god', 'mes', 'sum', 'data'])  # создание пустой заготовки итога
for i in ['PrTest2.csv',
          'PrTest1.csv'
          ]:
    dp = pd.read_csv(i,
                     # nrows=30
                     )
    dp = dp.fillna(0)
    tb = pd.pivot_table(dp, values=['sum', 'data'],
                        index=['schD2','schK2', 'muk', 'god', 'mes'],
                        aggfunc={'sum': sum, 'data': 'count'})
    tb.reset_index(inplace=True)
    tb = tb.fillna(0)
    pp = pp.append(tb)
dp = pp

tb = pd.pivot_table(dp, values=['sum', 'data'],
                    index=['schD2', 'schK2'],
                    columns=['muk'],
                    aggfunc={'sum': sum, 'data': sum})
tb.reset_index(inplace=True)
tb = tb.fillna(0)
tb['DK'] = tb.schD2.astype(str).str.cat(tb.schK2.astype(str), sep=';')

with open('prVn.pickle', 'rb') as f: prVn = pickle.load(f)
with open('pr.pickle', 'rb') as f: pr = pickle.load(f)
tb["vn"] = tb["DK"].map(prVn)
tb["opis"] = tb["DK"].map(pr)

tb = tb.fillna(0)

x = GetObject(None, "Excel.Application")
wb = x.Workbooks("TestPr.xlsx")
ws = wb.Worksheets("ГК")
StartRow = 6
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(tb.index) - 1,
                  StartCol + len(tb.columns) - 1)).Value = tb.values

# делаем помесячную ГК
tb = pd.pivot_table(dp, values=['sum', 'data'],
                    index=['schD2', 'schK2'],
                    columns=['god', 'mes'],
                    aggfunc={'sum': sum, 'data': sum}
                    , margins=True
                    )
tb.reset_index(inplace=True)
tb = tb.fillna(0)
tb['DK'] = tb.schD2.astype(str).str.cat(tb.schK2.astype(str), sep=';')

x = GetObject(None, "Excel.Application")
wb = x.Workbooks("TestPr.xlsx")
ws = wb.Worksheets("ГКмес")
StartRow = 6
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(tb.index) - 1,
                  StartCol + len(tb.columns) - 1)).Value = tb.values

tb.to_excel("output.xlsx")
# print(tb)
# print(dp.head(5))
# print(dp.columns)
# print(dp.shape)
# print(pp1.dtypes)
