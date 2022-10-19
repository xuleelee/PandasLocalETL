import pandas as pd
import numpy as np
import openpyxl as pxl
def my_gebi():
    milk = pd.read_csv('D:/TestForPandas/1.csv')
    information_table = pd.read_csv('D:/TestForPandas/2.csv')
    # print(df.head(5))

    # if milk['商品名称'].str.contains('↣')==True:
    sanjiaoku = milk.loc[milk['商品名称'].str.contains('△')]
    # print(type(data))
    print(sanjiaoku)
    print(sanjiaoku.empty)
    if sanjiaoku.empty:
        print('sanjiaoku is empty!')

    xinchangfa = milk.loc[milk['商品名称'].str.contains('✿')]
    if xinchangfa.empty:
        print('xinchangfa is empty!')

    else:
        print('xinchangfa is not empty!')
    # print(data.iloc[2]['订单号'])
    # print(data["B"])
    # df_p = data.pivot_table(index='商品名称',  # 透视的行，分组依据
    #                         values='商品数量',  # 值
    #                         aggfunc='sum'  # 聚合函数
    #                         )
    # # print(data)
    # df_p['商品数量'] = '共计' + df_p['商品数量'].astype('str') + '份'

    # print(df_p)

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
        data.to_excel(writer, '隔壁老王', index=True)
        # df_p.to_excel(writer, '隔壁海鲜', index=True)

        # Save the file
        writer.save()



my_gebi()