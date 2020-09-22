import numpy as np
import pandas as pd
from win32com.client import GetObject
import pickle
import matplotlib.pyplot as plt

pp = pd.DataFrame(columns=['DK', 'sum'])  # создание пустой заготовки итога
for i in ['PrTest1.csv',
          'PrTest2.csv'
          ]:
    dp = pd.read_csv(i,
                     # nrows=300
                     )
    dp = dp.fillna(0)
    dp = dp[['DK', 'sum', 'vn']]
    dp['abs'] = dp['sum'].abs()
    pp = pp.append(dp)
dp = pp
dp.to_csv('группы.csv', index=False)

# dp = pd.read_csv('группы.csv')
dp = dp.loc[dp['vn'] == 0]
k1 = 1000
k2 = 10000
k3 = 20000
k4 = 30000
k5 = 40000
k6 = 50000
k7 = 60000
k8 = 70000
k9 = 50000
dp.loc[(dp['abs'] > k9), 'mln'] = 10
dp.loc[dp['abs'] <= k1, 'mln'] = 1
dp.loc[(dp['abs'] > k1)&(dp['abs'] <= k2), 'mln'] = 2
# dp.loc[(dp['abs'] > k2)&(dp['abs'] <= k3), 'mln'] = 3
# dp.loc[(dp['abs'] > k3)&(dp['abs'] <= k4), 'mln'] = 4
# dp.loc[(dp['abs'] > k4)&(dp['abs'] <= k5), 'mln'] = 5
# dp.loc[(dp['abs'] > k5)&(dp['abs'] <= k6), 'mln'] = 6
# dp.loc[(dp['abs'] > k6)&(dp['abs'] <= k7), 'mln'] = 7
# dp.loc[(dp['abs'] > k7)&(dp['abs'] <= k8), 'mln'] = 8
# dp.loc[(dp['abs'] > k8)&(dp['abs'] <= k9), 'mln'] = 9

dp = dp.fillna(0)
tb = pd.pivot_table(dp, values=['abs', 'sum', 'vn'],
                    index=['mln'],
                    aggfunc={'abs': sum, 'sum': sum, 'vn': 'count'}
                    , margins=True
                    )
tb.reset_index(inplace=True)
tb = tb.fillna(0)

print(tb)

# tb.plot(x='mln', y='data')
# plt.grid(True)
# plt.show()

# tb.to_excel("output.xlsx")
# print(df[['abs']])
#
# print(dp.head(5))
# print(dp.columns)
# print(tb1.shape)
# print(tb.dtypes)
# dp1 = dp['rdate'].unique()
# dp1=pd.Series(dp1)
