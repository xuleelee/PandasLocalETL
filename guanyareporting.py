import  pandas as pd
import pandas_profiling

# read data
titanic = pd.read_excel('D:/12345.xlsx')
# 关键代码来了，就这一行代码
pandas_profiling.ProfileReport(titanic)
pfr = pandas_profiling.ProfileReport(titanic)
pfr.to_file('titanic.html')