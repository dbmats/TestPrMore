import pandas as pd
dp = pd.read_csv('PrTest.csv',
                   # nrows=30,
                   )
df = dp.iloc[:round(dp.shape[0]/2)]
df1 = dp.iloc[round(dp.shape[0]/2):]
df.to_csv('PrTest1.csv', index=False)
df1.to_csv('PrTest2.csv', index=False)

# собираем несколько файлов в один csv.

# df = pd.read_csv('PrTest.csv')
# dp = pd.read_csv('PrTest5.csv')
# df = df.append(dp)
# dp = pd.read_csv('PrTest3.csv')
# df = df.append(dp)
# dp = pd.read_csv('PrTest4.csv')
# df = df.append(dp)
# df.to_csv('PrTest.csv', index=False)

# df = pd.read_excel('1.xlsx', sheet_name='Лист1')
# dp = pd.read_excel('2.xlsx', sheet_name='Лист1')
# df = df.append(dp)
# dp = pd.read_excel('3.xlsx', sheet_name='Лист3')
# df = df.append(dp)
# dp = pd.read_excel('4.xlsx', sheet_name='Лист4')
# df = df.append(dp)
# dp = pd.read_excel('5.xlsx', sheet_name='Лист5')
# df = df.append(dp)
# df.to_csv('PrTest.csv', index=False)

# print(dp[['sum', 'kr1', 'kr']])
# print(dp)
# print(df.shape)
# print(df.dtypes)
