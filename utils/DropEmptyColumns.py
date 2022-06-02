
# importing libraries
import numpy as np
import pandas as pd

#/Users/cb-muneendra/Downloads/Exsalerate- Copy of Chargebee Migration Template (PC 2.0) - Subscriptions for customers.csv
# /Users/cb-muneendra/Downloads/Exsalerate- Copy of Chargebee Migration Template (PC 2.0) - Customers.csv
# dataFile = '/Users/cb-muneendra/Desktop/Foreup_customer_expected.xlsx'
# dataFile = '/Users/cb-muneendra/Desktop/Foreup_subscription_expected.xlsx'
dataFile = '/Users/cb-muneendra/Desktop/popmenu_subscription_expected.xlsx'


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


