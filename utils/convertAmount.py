import pandas as pd

file = "/Users/cb-muneendra/Git/cb_data_validation/readAPI/ds3files/Rewaatech_DS3_AllInvoices.xlsx"

df = pd.read_excel(file)
# df = pd.read_csv(file)

clist = ["invoice_total", "line_items_amount[0]", "line_items_item_level_discount_amount[0]", "line_items_tax_amount[0]", "invoice_tax", "payments_txn_amount[0]", "line_items_unit_amount[1]", "line_items_amount[1]", "line_items_item_level_discount_amount[1]", "line_items_tax_amount[1]"]

for col in clist:
    df[col] = df[col].div(100)
    # df[col] = df[col]*100

df.to_excel("updated_Rewaatech_DS3_AllInvoices.xlsx", index=False)
