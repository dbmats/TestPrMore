import pandas as pd
import win32com.client as win32

# выборка проводок со словами из списка

tt1 = pd.read_csv('Slova.csv')
tt = pd.Series(tt1['слова'].values, name='Value')
dd = pd.DataFrame(columns=['data', 'dok', 'org',
                           'schD', 'cD1', 'cD2', 'cD3',
                           'schK', 'cK1', 'cK2', 'cK3',
                           'sum', 'text', 'schD2', 'schK2', 'DK'])  # создание пустой заготовки
dd['slovo'] = str()

for i in ['PrTest1.csv',
          'PrTest2.csv'
          ]:
    dp = pd.read_csv(i,
                     # nrows=30
                     )
    dp = dp.fillna(0)

    dp = dp[['data', 'dok', 'org',
             'schD', 'cD1', 'cD2', 'cD3',
             'schK', 'cK1', 'cK2', 'cK3',
             'sum', 'text', 'schD2', 'schK2', 'DK']]
    dp['text'] = dp['text'].astype('str')
    dp['text'] = dp.text.str.lower()  # в строчные буквы

    for ind, val in enumerate(tt):
        dd1 = dp[dp['text'].str.contains(val)]
        dd1['slovo'] = (val)
        dd = dd.append(dd1)

dd['abs'] = dd['sum'].abs()
tb1 = pd.pivot_table(dd, values=['sum', 'data'],
                     index=['slovo'],
                     aggfunc={'sum': sum, 'data': 'count'})
tb1.reset_index(inplace=True)
tb1 = tb1.fillna(0)

dd.to_csv('НайденныеСлова.csv', index=False)
# dd = pd.read_csv('НайденныеСлова.csv')
tb = pd.pivot_table(dd, values=['sum', 'abs', 'data'],
                     index=['slovo', 'cD1', 'cD2', 'cK1', 'cK2', 'DK', 'text'],
                     aggfunc={'sum': sum, 'abs': sum, 'data': 'count'})
tb.reset_index(inplace=True)
tb = tb.fillna(0)
tb = tb[['slovo', 'cD1', 'cD2', 'cK1', 'cK2', 'DK', 'text', 'sum', 'abs', 'data']]

x = win32.GetObject(None, "Excel.Application")
wb = x.Workbooks("TestPr.xlsx")
ws = wb.Worksheets("Слова")
StartRow = 13
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(tb1.index) - 1,
                  StartCol + len(tb1.columns) - 1)).Value = tb1.values
ws = wb.Worksheets("СловаПР")
StartRow = 6
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(tb.index) - 1,
                  StartCol + len(tb.columns) - 1)).Value = tb.values

# tb.to_excel("output.xlsx")
# print(tb)
# print(dp)
# print(dp.head(15))
# print(tb.dtypes)
# print(tb.shape)
