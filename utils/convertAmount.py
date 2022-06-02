import pandas as pd

file = "/Users/cb-muneendra/Downloads/Inv_liv/Inv_liv.csv"

df = pd.read_csv(file)

clist = ["invoice[total]","line_items[amount][0]","payments[amount][0]","discounts[amount][0]","line_items[amount][1]","line_items[amount][2]","line_items[amount][3]","line_items[amount][4]"]

for col in clist:
    df[col] = df[col].div(100)

df.to_excel("output.xlsx", index=False)
