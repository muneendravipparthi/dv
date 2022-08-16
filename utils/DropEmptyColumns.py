# importing libraries
import pandas as pd

dataFile = '/Users/cb-muneendra/Desktop/Nsaas_Invoice_DS2.xlsx'

if dataFile.endswith('csv'):
    df = pd.read_csv(dataFile)
elif dataFile.endswith('xlsx'):
    df = pd.read_excel(dataFile)
columnsData = list(df.columns.values)
dic = df.count()
print(dic)

for col in columnsData:
    if dic.get(col) <= 0:
        df = df.drop(col, axis=1)
# print(list(df.head()))
# df.to_excel("Inv_liv1.xlsx", index=False)

print("========DS2==========")
for col in list(df.head()):
    print(col)

print("========DS3==========")
for col in list(df.head()):
    col = col.replace("[", "_")
    col = col.replace("]", "")
    col = col.replace("_0", "[0]")
    col = col.replace("_1", "[1]")
    col = col.replace("_2", "[2]")
    col = col.replace("_3", "[3]")
    col = col.replace("_4", "[4]")
    col = col.replace("_5", "[5]")
    col = col.replace("_6", "[6]")
    col = col.replace("_7", "[7]")
    col = col.replace("_8", "[8]")
    col = col.replace("_9", "[9]")
    print(col)
