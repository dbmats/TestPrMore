from win32com.client import GetObject
import pandas as pd

x = GetObject(None, "Excel.Application")
wb = x.Workbooks("TestPr.xlsx")
dp = pd.read_csv('выбранные.csv')

dp = dp[['data', 'dok', 'org',
         'schD', 'cD1', 'cD2', 'cD3',
         'schK', 'cK1', 'cK2', 'cK3',
         'sum', 'abs', 'text',
         'schD2', 'schK2', 'DK',
         'krr', 'vixodnoy', 'notch',
         'dubl', 'otritz', 'vir', 'redko']]


dp1 = dp.loc[dp['krr'] == 1]
ws = wb.Worksheets("1")
StartRow = 13
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(dp1.index) - 1,
                  StartCol + len(dp1.columns) - 1)).Value = dp1.values

dp['pp'] = dp['vixodnoy']+dp['notch']
dp2 = dp.loc[dp['pp'].isin([1, 2])]
ws = wb.Worksheets("2")
StartRow = 13
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(dp2.index) - 1,
                  StartCol + len(dp2.columns) - 1)).Value = dp2.values

dp3 = dp.loc[dp['dubl'] == 1]
ws = wb.Worksheets("3")
StartRow = 13
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(dp3.index) - 1,
                  StartCol + len(dp3.columns) - 1)).Value = dp3.values

dp4 = dp.loc[dp['dubl'] == 0]
dp4 = dp4.loc[dp4['otritz'] == 1]
ws = wb.Worksheets("4")
StartRow = 13
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(dp4.index) - 1,
                  StartCol + len(dp4.columns) - 1)).Value = dp4.values

dp5 = dp.loc[dp['krr'] == 0]
dp5 = dp5.loc[dp5['vixodnoy'] == 0]
dp5 = dp5.loc[dp5['notch'] == 0]
dp5 = dp5.loc[dp5['dubl'] == 0]
dp5 = dp5.loc[dp5['otritz'] == 0]
dp5 = dp5.loc[dp5['vir'] == 1]
ws = wb.Worksheets("5")
StartRow = 13
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(dp5.index) - 1,
                  StartCol + len(dp5.columns) - 1)).Value = dp5.values

dp6 = dp.loc[dp['redko'] == 1]
ws = wb.Worksheets("6")
StartRow = 13
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(dp6.index) - 1,
                  StartCol + len(dp6.columns) - 1)).Value = dp6.values

dp7 = dp.loc[dp['krr'] == 0]
dp7 = dp7.loc[dp7['vixodnoy'] == 0]
dp7 = dp7.loc[dp7['notch'] == 0]
dp7 = dp7.loc[dp7['dubl'] == 0]
dp7 = dp7.loc[dp7['otritz'] == 0]
dp7 = dp7.loc[dp7['vir'] == 0]
dp7 = dp7.loc[dp7['redko'] == 0]
ws = wb.Worksheets("7")
StartRow = 13
StartCol = 1
ws.Range(ws.Cells(StartRow, StartCol),
         ws.Cells(StartRow + len(dp7.index) - 1,
                  StartCol + len(dp7.columns) - 1)).Value = dp7.values