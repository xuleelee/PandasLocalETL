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
import openpyxl as pxl

milk = pd.read_csv('D:/TestForPandas/1.csv')
information_table = pd.read_csv('D:/TestForPandas/2.csv')
# print(df.head(5))
data = milk.loc[milk['商品名称'].str.contains('❄')]
df_p = data.pivot_table(index='商品名称',    # 透视的行，分组依据
                      values='商品数量',    # 值
                      aggfunc='sum'    # 聚合函数
                     )
# print(data)
df_p['商品数量'] = '共计' + df_p['商品数量'].astype('str') + '份'


print(df_p)

filename = '最终的统计后.xlsx'
excel_book = pxl.load_workbook(filename)

with pd.ExcelWriter(filename, engine='openpyxl') as writer:
    # Your loaded workbook is set as the "base of work"
    writer.book = excel_book

    # Loop through the existing worksheets in the workbook and map each title to\
    # the corresponding worksheet (that is, a dictionary where the keys are the\
    # existing worksheets' names and the values are the actual worksheets)
    writer.sheets = {worksheet.title: worksheet for worksheet in excel_book.worksheets}



    # Write the new data to the file without overwriting what already exists
    df_p.to_excel(writer, '牛奶', index=True)

    # Save the file
    writer.save()

