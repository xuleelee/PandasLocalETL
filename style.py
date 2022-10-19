import xlwt
import xlrd
from xlutils.copy import copy
#创建execl

#单元格上色
def color_execl(file_name):
    styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour red;')  # 红色
    rb = xlrd.open_workbook(file_name)      #打开t.xls文件
    ro = rb.sheets()[0]                     #读取表单0
    wb = copy(rb)                           #利用xlutils.copy下的copy函数复制
    ws = wb.get_sheet(0)                    #获取表单0
    col = 0                                 #指定修改的列
    for i in range(ro.nrows):               #循环所有的行
        result = int(ro.cell(i, col).value)
        print(result)
        if result == 2:                     #判断是否等于2
            ws.write(i,col,ro.cell(i, col).value,styleBlueBkg)
    wb.save(file_name)

if __name__ == '__main__':
    file_name = '最终的统计后.xls'

    color_execl(file_name)
    # color_execl(file_name)