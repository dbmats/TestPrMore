from win32com.client import GetObject
import pandas as pd
import pickle

# если:
# 1 сверка сходится;
# 2 кол-во проводок больше МУК не намного более 1т.шт., или меньше;
# 3 не нужны корректировки словаря со внутренностями.
# то здесь добавляем в исходную совок. ДК и внутр-сть и делаем выборку

# выбрал и сохранил выбранные
# pp = pd.DataFrame(columns=['data', 'dok', 'org', 'schD', 'cD1', 'cD2', 'cD3',
#                            'schK', 'cK1', 'cK2', 'cK3', 'sum', 'text', 'time',
#                            'god', 'mes', 'denb', 'schD2', 'schK2', 'muk', 'DK', 'vn',
#                            'xxx', 'xx', 'x1', 'x2', 'vn'])  # создание пустой заготовки итога
# for i in ['PrTest2.csv',
#           'PrTest1.csv'
#           ]:
#     dp = pd.read_csv(i,)
#     dp = dp.fillna(0)

#     dp = dp.loc[dp['muk'] == True]
#     dp = dp.loc[dp['vn'] == 0]
#     pp = pp.append(dp)
# dp = pp
# dp.to_csv('выбранные.csv', index=False)

# dp = pd.read_csv('выбранные.csv')

# добавляем количества по проводкам и меньше 12 раз в год
with open('redko.pickle', 'rb') as f: redko = pickle.load(f)
dp["kolPr"] = dp["DK"].map(redko)
dp['redko'] = (dp['kolPr'] < 12).astype(int)

# # добавляем круглость, одинаковые суммы, отрицательные суммы, выручка
# dp['tzel'] = (((dp['sum'].astype('str')).str.split('.').str.get(1)) == '0') \
#     .astype(int)
# dp['kr1'] = ((((dp['sum']/1000).astype('str')).str.split('.').str.get(1))).str[0:3]
# dp['kr'] = ((dp['kr1'] == '0')|(dp['kr1'] == '999')|(dp['kr1'] == '888')|(dp['kr1'] == '777')
#             |(dp['kr1'] == '666')|(dp['kr1'] == '555')|(dp['kr1'] == '444')|(dp['kr1'] == '333')
#             |(dp['kr1'] == '222')|(dp['kr1'] == '111')).astype('int')
# dp['krr'] = dp['kr']*dp['tzel']
# dp['dubl'] = dp['sum'].abs().duplicated(False).astype(int)
# dp['otritz'] = (dp['sum'] < 0).astype(int)
# dp['vir'] = (dp['DK'] == '62;90').astype(int)
# dp['abs'] = dp['sum'].abs()

# добавляем выходные и ночь
# with open('vixDen.pickle', 'rb') as f: vixDen = pickle.load(f)
# dp['DT'] = dp.god.astype(str)+'#'+dp.mes.astype(str)+'#'+dp.denb.astype(str)
# dp['vixodnoy'] = dp['DT'].map(vixDen)
# dp['notch'] = ((dp['time'] >= 1)&(dp['time'] <= 5)).astype(int)
# dp = dp.fillna(0)

dp.to_csv('выбранные.csv', index=False)

# переносим выбранные в РД
# dp = dp[['data', 'dok', 'org',
#          'schD', 'cD1', 'cD2', 'cD3',
#          'schK', 'cK1', 'cK2', 'cK3',
#          'sum', 'abs', 'text',
#          'schD2', 'schK2', 'DK',
#          'krr', 'vixodnoy', 'notch',
#          'dubl', 'otritz', 'vir', 'redko']]
#
# x = GetObject(None, "Excel.Application")
# wb = x.Workbooks("TestPr.xlsx")
# ws = wb.Worksheets("выбранные")
# StartRow = 9
# StartCol = 1
# ws.Range(ws.Cells(StartRow, StartCol),
#          ws.Cells(StartRow + len(dp.index) - 1,
#                   StartCol + len(dp.columns) - 1)).Value = dp.values

# dp.to_excel("output1.xlsx")
# print(dp[['DT']])
# print(dp)
# print(dp.shape)
# print(dp.dtypes)
