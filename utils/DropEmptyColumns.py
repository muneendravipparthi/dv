
# importing libraries
import numpy as np
import pandas as pd


dataFile = '/Users/cb-muneendra/Desktop/leeto_Live_subscription_expected.xlsx'


if dataFile.endswith('csv'):
    df = pd.read_csv(dataFile)
elif dataFile.endswith('xlsx'):
    df = pd.read_excel(dataFile)
columnsData = list(df.columns.values)
dic = df.count()
print(dic)

for col in columnsData:
    if dic.get(col) <= 0:
        df = df.drop(col,axis = 1)
# print(list(df.head()))
# df.to_excel("Inv_liv1.xlsx", index=False)
print("===================")
for col in list(df.head()):
    print(col)

for col in list(df.head()):
    col = col.replace("[","_")
    col = col.replace("]","")
    col = col.replace("_0","[0]")
    col = col.replace("_1","[1]")
    col = col.replace("_2","[2]")
    print(col)
