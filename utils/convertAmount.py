import pandas as pd

file = "/Users/cb-muneendra/Downloads/Beekeeper Live - beekeeperlive-invoice full.csv"

df = pd.read_csv(file)

clist = ["invoice[total]", "payment[amount]", "taxes[amount][0]", "line_items[unit_amount][0]", "line_items[amount][0]", "line_items[unit_amount][1]", "line_items[amount][1]", "line_items[unit_amount][2]", "line_items[amount][2]", "line_items[unit_amount][3]", "line_items[amount][3]", "line_items[tax1_amount][0]", "line_items[tax2_amount][0]"]

for col in clist:
    # df[col] = df[col].div(100)
    df[col] = df[col]*100

df.to_excel("beekeeperlive_Inv_DS2.xlsx", index=False)
