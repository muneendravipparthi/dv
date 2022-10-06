import pandas as pd

file = "/Users/cb-muneendra/Downloads/Rewaatech - Chargebee Migration Template (PC 2.0) - invoices-phase6.csv"

# df = pd.read_excel(file)
df = pd.read_csv(file)

clist = ["invoice[total]", "line_items[amount][0]", "line_items[item_level_discount1_amount][0]", "line_items[tax1_amount][0]", "taxes[amount][0]", "payments[amount][0]", "line_items[unit_amount][1]", "line_items[amount][1]", "line_items[item_level_discount1_amount][1]", "line_items[tax1_amount][1]"]

for col in clist:
    # df[col] = df[col].div(100)
    df[col] = df[col]*100

df.to_excel("Chargebee Migration Template (PC 2.0) - invoices-phase6.xlsx", index=False)
