import pandas as pd

# source_path = "/Users/cb-muneendra/Downloads/invoices_pixellu-test_09_May_2022_13_14_21/LineItems.csv"
#
# for i,chunk in enumerate(pd.read_csv(source_path, chunksize=50000)):
#     chunk.to_csv('pixellulineitemfile{}.csv'.format(i), index=False)
#     print('pixellulineitemfile{}.csv'.format(i))


source_path = "/Users/cb-muneendra/Downloads/Inv_liv/Inv_liv.csv"

df = pd.read_csv(source_path)
clist = ["invoice[total]","line_items[amount][0]","payments[amount][0]","discounts[amount][0]","line_items[amount][1]","line_items[amount][2]","line_items[amount][3]","line_items[amount][4]"]

for col in clist:
    df[col] = df[col].div(100)
# df = pd.read_excel(source_path)
print("started spliting")
df1 = df.iloc[:50000]
df1.to_excel('pixellu_Live_invoice1.xlsx', index=False)
df2 = df.iloc[50001:100000]
df2.to_excel('pixellu_Live_invoice2.xlsx', index=False)
df3 = df.iloc[100001:150000]
df3.to_excel('pixellu_Live_invoice3.xlsx', index=False)
df4 = df.iloc[150001:200000]
df4.to_excel('pixellu_Live_invoice4.xlsx', index=False)
df5 = df.iloc[200001:250000]
df5.to_excel('pixellu_Live_invoice5.xlsx', index=False)
df6 = df.iloc[250001:300000]
df6.to_excel('pixellu_Live_invoice6.xlsx', index=False)
df7 = df.iloc[300001:350000]
df7.to_excel('pixellu_Live_invoice7.xlsx', index=False)
df8 = df.iloc[350001:400000]
df8.to_excel('pixellu_Live_invoice8.xlsx', index=False)
df9 = df.iloc[400001:450000]
df9.to_excel('pixellu_Live_invoice9.xlsx', index=False)
df10 = df.iloc[450001:500000]
df10.to_excel('pixellu_Live_invoice10.xlsx', index=False)
df11 = df.iloc[500001:550000]
df11.to_excel('pixellu_Live_invoice11.xlsx', index=False)
df12 = df.iloc[550001:600000]
df12.to_excel('pixellu_Live_invoice12.xlsx', index=False)
df13 = df.iloc[600001:650000]
df13.to_excel('pixellu_Live_invoice13.xlsx', index=False)
df14 = df.iloc[650001:]
df14.to_excel('pixellu_Live_invoice14.xlsx', index=False)
print("ended")









