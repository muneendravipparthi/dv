import pandas as pd


def sorting_strings(list):
    list.sort()
    strvalue = ' ; '.join(map(str, list))
    return strvalue.lower()


# df = pd.read_excel('/Users/cb-muneendra/Desktop/SampleData.xlsx')
# df['clubbed'] = df.apply(lambda x: '%s_%s_%s_%s' % (x['col1'], x['col2'], x['col3'], x['col4']), axis=1)
#
# df['clubbed1'] = df.apply(lambda x: '%s' % ([x['col1'], x['col2'], x['col3'], x['col4']]), axis=1)
# df['clubbed2'] = df.apply(lambda x: '%s' % (sorting_strings([x['col1'], x['col2'], x['col3'], x['col4']])), axis=1)
# print(df)

df = pd.read_excel('/Users/cb-muneendra/Git/cb_data_validation/readAPI/ds3files/microshareltd_us_DS3_AllSubscriptions.xlsx')
df = df[["subscription_id", "item_item_price_id[0]", "item_item_price_id[1]", "item_item_price_id[2]",
         "item_item_price_id[3]", "item_item_price_id[4]", "item_item_price_id[5]"]]

df['item_item_price_id'] = df.apply(lambda x: '%s' % (sorting_strings(
    [str(x['item_item_price_id[0]']), str(x['item_item_price_id[1]']),
     str(x['item_item_price_id[2]']), str(x['item_item_price_id[3]']),
     str(x['item_item_price_id[4]']), str(x['item_item_price_id[5]']),
     ])), axis=1)
df1 = df[["subscription_id", "item_item_price_id"]]
df1.to_excel("headrushtech_subscriptions_expected2.xlsx", index=False)
listcoltomerge = ["item_item_price_id[0]", "item_item_price_id[1]", "item_item_price_id[2]",
         "item_item_price_id[3]", "item_item_price_id[4]", "item_item_price_id[5]"]
df2 = pd.DataFrame
for col in listcoltomerge:
    df[col] = df[col].astype(str)
df['merge'] = df[listcoltomerge].values.tolist()
print(df['merge'])
df['merge1'] = df['merge'].apply(lambda x: sorting_strings(x))
print(df['merge1'])
df.to_excel("microshareltd_uk_DS3_AllSubscriptions3.xlsx", index=False)
