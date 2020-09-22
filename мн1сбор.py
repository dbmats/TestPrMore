import pandas as pd
import pickle

# дополнение совокупности столбцами и формирование csv
dp = pd.read_excel('МНрабочая.xlsx', sheet_name='1',
                   # nrows=300,
                   )

# если дата в формате (день.мес.год час:мин:сек)
dp['time'] = (((dp['data'].astype('str')).str.split(' ').str.get(1)).str.split(':').str.get(0)).astype('int64')
dp['god'] = ((dp['data'].astype('str')).str.split(' ').str.get(0)).str.split('.').str.get(2)
dp['mes'] = ((dp['data'].astype('str')).str.split(' ').str.get(0)).str.split('.').str.get(1)
dp['denb'] = ((dp['data'].astype('str')).str.split(' ').str.get(0)).str.split('.').str.get(0)

# если дата в формате (год.мес.день)
# dp['god'] = (dp['data'].astype('str')).str.split('-').str.get(0)
# dp['mes'] = (dp['data'].astype('str')).str.split('-').str.get(1)
# dp['denb'] = (dp['data'].astype('str')).str.split('-').str.get(2)
# dp['time'] = 0

dp['schD2'] = (dp['schD'].astype('str')).str[0:2]
dp['schK2'] = (dp['schK'].astype('str')).str[0:2]
dp['DK'] = dp.schD2.astype(str).str.cat(dp.schK2.astype(str), sep=';')

with open('prVn.pickle', 'rb') as f: prVn = pickle.load(f) # добавить
dp["vn"] = dp["DK"].map(prVn)                               # добавить

dp['xxx'] = (dp['sum'].abs() * 1000000).astype('str')
dp['xx'] = dp.xxx.str[0:2]
dp['x1'] = dp.xxx.str[0:1]
dp['x2'] = dp.xx.str[1:2]

dp.to_csv('PrTest2.csv', index=False)

# dp.to_excel("output.xlsx")
# print(tb1)
# print(dp.head(5))
# print(dp.columns)
# print(dp.shape)
# print(dp[['schD', 'schD2']])
# print(dp.dtypes)
