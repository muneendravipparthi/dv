
# importing libraries
import numpy as np
import pandas as pd


dataFile = '/Users/cb-muneendra/Downloads/output_22_ xpendy_data_1 2/invoice_output6936678840555499885/invoice_cb_template.csv'


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


