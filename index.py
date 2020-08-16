import pandas as pd
import numpy as np

# df = pd.read_excel(r'./week-rank.xls')
# print(df)

# data = np.arange(1,101).reshape((10,10))
data = [[1,2,3,4,5,6,7,8,9,0],[1,2,3,4,5,6,7,8,9,0]]
print(data)
data_df = pd.DataFrame(data)
data_df.columns = ['A','B','C','D','E','F','G','H','I','J']
# data_df.index = ['a','b','c','d','e','f','g','h','i','j']

writer = pd.ExcelWriter('my.xlsx')
data_df.to_excel(writer,float_format='%.5f')
writer.save()