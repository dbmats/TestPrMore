import re
import pandas as pd
import win32com.client as win32

# анализ слов при анализе выручки, К60,76, оплате поставщикам

dd = pd.DataFrame(columns=['text', 'DK'])  # создание пустой заготовки
for i in ['PrTest1.csv',
          'PrTest2.csv'
          ]:
    dp = pd.read_csv(i,
                     # nrows=30
                     )
    dp = dp.fillna(0)
    dp = dp[['text', 'DK']]
    dp = dp.loc[dp['DK'].isin(['62;90', '20;60', '23;60', '25;60', '26;60', '44;60',
                '20;76', '23;76', '25;76', '26;76', '44;76', '32;60', '32;76',
                '60;51', '60;52', '76;51', '76;52'])]
    dp['text'] = dp['text'].astype('str')
    dd = dd.append(dp)
dp = dd

dp1 = dp.loc[dp['DK'] == '62;90']
dp1 = pd.Series(dp1['text'].values, name='Value')
dp1 = dp1.str.cat(sep=' ').lower()
dd = pd.DataFrame(re.findall(r'\b[а-я]{3,35}\b', dp1))
dd['x'] = 1
tb = pd.pivot_table(dd, values=['x'],
                    index=0,
                    aggfunc={'x': sum}).sort_values(by=['x'], ascending=False)
tb.reset_index(inplace=True)

dp2 = dp.loc[dp['DK'].isin(['20;60', '23;60', '25;60', '26;60', '44;60',
                            '20;76', '23;76', '25;76', '26;76', '44;76', '32;60', '32;76'])]
dp1 = pd.Series(dp2['text'].values, name='Value')
dp1 = dp1.str.cat(sep=' ').lower()
dd = pd.DataFrame(re.findall(r'\b[а-я]{3,35}\b', dp1))
dd['x'] = 1
tb1 = pd.pivot_table(dd, values=['x'],
                     index=0,
                     aggfunc={'x': sum}).sort_values(by=['x'], ascending=False)
tb1.reset_index(inplace=True)

dp3 = dp.loc[dp['DK'].isin(['60;51', '60;52', '76;51', '76;52'])]
dp1 = pd.Series(dp3['text'].values, name='Value')
dp1 = dp1.str.cat(sep=' ').lower()
dd = pd.DataFrame(re.findall(r'\b[а-я]{3,35}\b', dp1))
dd['x'] = 1
tb2 = pd.pivot_table(dd, values=['x'],
                     index=0,
                     aggfunc={'x': sum}).sort_values(by=['x'], ascending=False)
tb2.reset_index(inplace=True)

x = win32.GetObject(None, "Excel.Application")
wb = x.Workbooks("TestPr.xlsx")
ws = wb.Worksheets("слова1")
StartRow = 14
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(tb.index) - 1,
                  StartCol + len(tb.columns) - 1)).Value = tb.values
StartRow = 14
StartCol = 4
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(tb1.index) - 1,
                  StartCol + len(tb1.columns) - 1)).Value = tb1.values
StartRow = 14
StartCol = 7
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(tb2.index) - 1,
                  StartCol + len(tb2.columns) - 1)).Value = tb2.values

# print(tb2)
# print(dd[['0']])
# print(dp)
# print(dd.dtypes)
