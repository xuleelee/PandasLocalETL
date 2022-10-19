import pandas as pd
import numpy as np
import openpyxl as pxl

import pandas as pd
import numpy as np
import openpyxl as pxl

list = []
list_not_order = []

def highlight_number(number):
    criteria = number == '35435'
    print(['background-color:yellow' if i else '' for i in criteria])
    return ['background-color:yellow' if i else '' for i in criteria]

def my_mianbao():




    original_table = pd.read_csv('D:/TestForPandas/1.csv')
    information_table = pd.read_excel('D:/TestForPandas/mianbao.xlsx', sheet_name='Sheet1')



    result = original_table.merge(information_table, on="商品名称")
    mianbao = result[['订单号','商品名称','商品数量']].copy()

    # mianbao.sort_values("订单号", ascending=[False],inplace=True)
    mianbao.sort_values("订单号",  inplace=True)
    mianbao['订单号']= mianbao['订单号'].astype(str)
    mianbao['订单号'] = mianbao['订单号'].str[13:18]
    mianbao[['商品名称', '商品数量']] = mianbao[['商品数量','商品名称']]
    # print(mianbao['订单号'])
    # print(mianbao['订单号'])
    chongfu = mianbao['订单号']
    baba = mianbao.copy()
    print(baba)

    # rows_series = mianbao[['订单号']].duplicated(keep=False)
    # print(rows_series)
    # rows = rows_series[rows_series].index.values
    # print(rows)
    # yanse = mianbao.style.apply(lambda x: ['background: yellow' if x.name in rows else '' for i in x], axis=1)
    # yanse = mianbao.style.apply(lambda x: ['background: yellow' if x.name in rows else '' for i in x])
    # mianbao.style.format({'订单号':'{:.2%}'})
    # # print(mianbao['Duplicate'])
    # print(mianbao)
    if not result.empty:
        vege_pivot = result.pivot_table(index='商品名称',  # 透视的行，分组依据
                                        values='商品数量',  # 值
                                        aggfunc='sum',  # 聚合函数
                                        margins=True,
                                        margins_name='合计'
                                        )


        vege_pivot['商品数量'] = '共计' + vege_pivot['商品数量'].astype('str') + '份'
        baba.style.apply(highlight_number)

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
            vege_pivot.to_excel(writer, '面包总数', index=True)
            baba.to_excel(writer, '宝宝宝宝', index=True)
            # Save the file
            writer.save()
            list.append('面包')
    else:
        list_not_order.append('面包')


my_mianbao()