import pandas as pd
import numpy as np
import os
import xlrd
 
# 读取第一个文件
df1 = pd.read_excel('C:\Users\xin\Desktop\py-stu\1.xlsx')
# 读取第二个文件
df2 = pd.read_excel('C:\Users\xin\Desktop\py-stu\2.xlsx')
# 按行合并两个文件
result = pd.concat([df1, df2])
# 将结果保存到新的Excel文件
result.to_excel('C:\Users\xin\Desktop\py-stu\hang.xlsx', index=False)