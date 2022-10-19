# # This is a sample Python script.
#
# # Press Shift+F10 to execute it or replace it with your code.
# # Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
#
#
# def print_hi(name):
#     # Use a breakpoint in the code line below to debug your script.
#     print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.
#
#
# # Press the green button in the gutter to run the script.
# if __name__ == '__main__':
#     print_hi('PyCharm')
#
# # See PyCharm help at https://www.jetbrains.com/help/pycharm/

import pandas as pd
import numpy as np



original_table = pd.read_csv('D:/TestForPandas/1.csv')
information_table = pd.read_excel('D:/TestForPandas/information.xlsx',sheet_name='vege')

original_table.loc[(original_table['商品名称'].str.contains('上海青') ),['商品数量']] = original_table.loc[(original_table['商品名称'].str.contains('上海青') ),['商品数量']] *3
original_table.loc[(original_table['商品名称'].str.contains('青菜心') ),['商品数量']] = original_table.loc[(original_table['商品名称'].str.contains('青菜心') ),['商品数量']] *2
original_table.loc[(original_table['商品名称'].str.contains('芥蓝') ),['商品数量']] = original_table.loc[(original_table['商品名称'].str.contains('芥蓝') ),['商品数量']] *2
original_table.loc[(original_table['商品名称'].str.contains('樱桃萝卜') ),['商品数量']] = original_table.loc[(original_table['商品名称'].str.contains('樱桃萝卜') ),['商品数量']] *3

print(original_table)
result = original_table.merge(information_table,on="商品名称")
# result = pd.merge(original_table,information_table.loc[:,['商品名称','英语名字']],how='right',on = '商品名称')
# print(result)

vege_pivot = result.pivot_table(index='english',    # 透视的行，分组依据
                      values='商品数量',    # 值
                      aggfunc='sum'    # 聚合函数
                     )
print(vege_pivot)

final_df = result = vege_pivot.merge(information_table,on="english")

print(final_df)

# without_duplication = final_df.drop_duplicates(['english'],keep='last',inplace=True)
final_test=final_df.duplicated(subset=['english'],keep='last')
print(final_test)

finally_test=final_df.drop_duplicates(subset=['english'],keep='last')

print(finally_test)






writer = pd.ExcelWriter('Vlookup统计后.xlsx')
vege_pivot.to_excel(writer,sheet_name='蔬菜')
final_df.to_excel(writer,sheet_name='sheet2')
finally_test.to_excel(writer,sheet_name='sheet4')
writer.save()