import pandas as pd

# file = "/Users/cb-muneendra/Downloads/invoices.csv"
file = "/Users/cb-muneendra/Downloads/invoices (1).csv"

df = pd.read_csv(file)

col1 = ["invoice[id]", "invoice[stripe_id]", "invoice[customer_id]", "invoice[subscription_id]", "line_items[amount][0]", "line_items[date_from][0]", "line_items[date_to][0]", "line_items[description][0]", "old_line_items[entity_id][0]", "line_items[entity_id][0]", "line_items[entity_type][0]", "line_items[id][0]", "line_items[quantity][0]", "old_invoice[customer_id]"]
col2 = ["invoice[id]", "invoice[stripe_id]", "invoice[customer_id]", "invoice[subscription_id]", "line_items[amount][1]", "line_items[date_from][1]", "line_items[date_to][1]", "line_items[description][1]", "old_line_items[entity_id][1]", "line_items[entity_id][1]", "line_items[entity_type][1]", "line_items[id][1]", "line_items[quantity][1]", "old_invoice[customer_id]"]
col3 = ["invoice[id]", "invoice[stripe_id]", "invoice[customer_id]", "invoice[subscription_id]", "line_items[amount][2]", "line_items[date_from][2]", "line_items[date_to][2]", "line_items[description][2]", "old_line_items[entity_id][2]", "line_items[entity_id][2]", "line_items[entity_type][2]", "line_items[id][2]", "line_items[quantity][2]", "old_invoice[customer_id]"]
col4 = ["invoice[id]", "invoice[stripe_id]", "invoice[customer_id]", "invoice[subscription_id]", "line_items[amount][3]", "line_items[date_from][3]", "line_items[date_to][3]", "line_items[description][3]", "line_items[entity_id][3]", "old_line_items[entity_id][3]", "line_items[entity_type][3]", "line_items[id][3]", "line_items[quantity][3]", "old_invoice[customer_id]"]
col5 = ["invoice[id]", "invoice[stripe_id]", "invoice[customer_id]", "invoice[subscription_id]", "line_items[amount][4]", "line_items[date_from][4]", "line_items[date_to][4]", "line_items[description][4]", "old_line_items[entity_id][4]", "line_items[entity_id][4]", "line_items[entity_type][4]", "line_items[id][4]", "line_items[quantity][4]", "old_invoice[customer_id]"]

headorg = ["invoice[id]", "invoice[stripe_id]", "invoice[customer_id]", "invoice[subscription_id]", "line_items[amount]", "line_items[date_from]", "line_items[date_to]", "line_items[description]", "old_line_items[entity_id]", "line_items[entity_id]", "line_items[entity_type]", "line_items[id]", "line_items[quantity]", "old_invoice[customer_id]"]

df1 = df[col1]
df2 = df[col2]
df3 = df[col3]
df4 = df[col4]
df5 = df[col5]
newheaders = {}
for i in range(0, 13):
   newheaders[col1[i]] =  headorg[i]
   df1.rename(columns=newheaders, inplace=True)
   newheaders = {}
   newheaders[col2[i]] = headorg[i]
   df2.rename(columns=newheaders, inplace=True)
   newheaders = {}
   newheaders[col3[i]] = headorg[i]
   df3.rename(columns=newheaders, inplace=True)
   newheaders = {}
   newheaders[col4[i]] = headorg[i]
   df4.rename(columns=newheaders, inplace=True)
   newheaders = {}
   newheaders[col5[i]] = headorg[i]
   df5.rename(columns=newheaders, inplace=True)
df1 = df1[df1['line_items[entity_id]'].notnull()]
df2 = df2[df2['line_items[entity_id]'].notnull()]
df3 = df3[df3['line_items[entity_id]'].notnull()]
df4 = df2[df4['line_items[entity_id]'].notnull()]
df5 = df2[df5['line_items[entity_id]'].notnull()]
df_new = df1.append(df2)
df_new = df_new.append(df3)
df_new = df_new.append(df4)
df_new = df_new.append(df5)

# df_new = df_new['line_items[entity_id]'].replace('', np.nan, inplace=True)
# df_new = df_new.dropna(subset=['line_items[entity_id]'], inplace=True)
df_new.to_excel("lineitemsconverted_new1.xlsx", index=False)
# df2.to_excel("lineitemsconverted.xlsx", index=False)
# df2 = df2[df2['line_items[entity_id]'].notnull()]
# df2.to_excel("lineitemsconverted1.xlsx", index=False)