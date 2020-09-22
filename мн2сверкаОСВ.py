import win32com.client as win32
import pandas as pd
import pickle

# формирование Дт Кт обороты для сверки с ОСВ
pp1 = pd.DataFrame(columns=['DK', 'mes'])
pp = pd.DataFrame(columns=['schD2','schK2', 'god', 'sum'])  # создание пустой заготовки итога

for i in ['PrTest2.csv',
          'PrTest1.csv'
          ]:
    dp = pd.read_csv(i,
                     # nrows=30
                     )
    dp = dp.fillna(0)
    tb = pd.pivot_table(dp, values=['sum'],
                        index=['schD2','schK2', 'god'],
                        aggfunc={'sum': sum})
    tb.reset_index(inplace=True)
    tb = tb.fillna(0)
    pp = pp.append(tb)

    tb1 = pd.pivot_table(dp, values=['mes'],
                        index=['DK'],
                        aggfunc={'mes': 'count'})
    tb1.reset_index(inplace=True)
    tb1 = tb1.fillna(0)
    pp1 = pp1.append(tb1)

# создал словарь с количествами по проводкам
tb = pd.pivot_table(pp1, values=['mes'],
                        index=['DK'],
                        aggfunc={'mes': sum})
tb.reset_index(inplace=True)
tb = tb.fillna(0)
redko = dict(sorted(tb.values.tolist()))
with open('redko.pickle', 'wb') as f: pickle.dump(redko, f)

dp = pp
tb = pd.pivot_table(dp, values=['sum'],
                    index=['schD2', 'god'],
                    aggfunc={'sum': sum})
tb['sumK'] = 0.0
tb.columns = ['sumD', 'sumK']

tb1 = pd.pivot_table(dp, values=['sum'],
                     index=['schK2', 'god'],
                     aggfunc={'sum': sum})
tb1['sumD'] = 0.0
tb1 = tb1[['sumD', 'sum']]
tb1.columns = ['sumD', 'sumK']

fr = [tb, tb1]
pp = pd.concat(fr)
pp.reset_index(inplace=True)

pp1 = pd.pivot_table(pp, values=['sumD', 'sumK'],
                     columns=['god'],
                     index=['schD2'],
                     aggfunc={'sumD': sum, 'sumK': sum})
pp1.reset_index(inplace=True)
pp1 = pp1.fillna(0)

x = win32.GetObject(None, "Excel.Application")
wb = x.Workbooks("TestPr.xlsx")
ws = wb.Worksheets("сверка")
StartRow = 8
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(pp1.index) - 1,
                  StartCol + len(pp1.columns) - 1)).Value = pp1.values

# pp.to_excel("output.xlsx")
# print(tb)
# print(dp.head(5))
# print(tb.columns)
# print(pp1.shape)
# print(tb.dtypes)
