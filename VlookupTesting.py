import pandas as pd
import numpy as np



information_table = pd.read_excel('D:/TestForPandas/informationtest.xlsx',sheet_name='vege')

# information_table.loc[(information_table['A'] .str.contains('Spring Onions') ),['B']] = information_table.loc[(information_table['A'] == ''),['B']]*3

# if information_table.loc[(information_table['A'] .str.contains('Spring Onions') ):
information_table.loc[(information_table['A'] .str.contains('abc Onions') ),['B']] =information_table.loc[(information_table['A'] .str.contains('Spring Onions') ),['B']]*33

print(information_table)

# result = pd.merge(original_table,information_table.loc[:,['商品名称','英语名字']],how='right',on = '商品名称')
# print(result)

vege_pivot = information_table.pivot_table(index='A',    # 透视的行，分组依据
                      values='B',    # 值
                      aggfunc='sum'    # 聚合函数
                     )
print(vege_pivot)

# vege_pivot.loc[(vege_pivot.A == 'eggplant'),'B'] = "1000"

# mask = (vege_pivot['A'] =='eggplant')
# print(mask)
# # .loc[] 赋值
# vege_pivot.loc[mask, 'B'] = 1000
#
#
# print(vege_pivot)
