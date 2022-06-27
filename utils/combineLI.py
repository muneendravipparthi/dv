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

df = pd.read_excel('/Users/cb-muneendra/Desktop/foreup_subscription_expeted2.xlsx')
df = df[["customer[id]", "subscription[id]", "subscription[status]", "subscription[started_at]", "subscription[current_term_start]", "subscription[current_term_end]", "subscription_items[item_price_id][0]", "subscription_items[item_price_id][1]", "subscription_items[quantity][1]", "subscription_items[unit_price][1]", "subscription_items[item_price_id][2]", "subscription_items[quantity][2]", "subscription_items[unit_price][2]", "subscription_items[item_price_id][3]", "subscription_items[quantity][3]", "subscription_items[unit_price][3]", "subscription_items[item_price_id][4]", "subscription_items[quantity][4]", "subscription_items[unit_price][4]", "subscription_items[item_price_id][5]", "subscription_items[quantity][5]", "subscription_items[unit_price][5]", "subscription_items[item_price_id][6]", "subscription_items[quantity][6]", "subscription_items[unit_price][6]", "subscription_items[item_price_id][7]", "subscription_items[quantity][7]", "subscription_items[unit_price][7]", "subscription_items[item_price_id][8]", "subscription_items[quantity][8]", "subscription_items[unit_price][8]", "subscription_items[item_price_id][9]", "subscription_items[quantity][9]", "subscription_items[unit_price][9]", "coupon_ids[0]"]]
for col in list(df.head()):
    try:
        df[col] = df[col].str.lower()
        df[col] = df[col].astype(str).apply(lambda x: x.replace('.0', ''))
    except Exception as e:
        print("Exception for column : {} and exception is : {}".format(col, e))

# df['Subscription_Item_1'] = df[['subscription_items[item_price_id][1]', 'subscription_items[quantity][1]', 'subscription_items[unit_price][1]']].agg('-'.join, axis=1)

df['Subscription_Item_0'] = df['subscription_items[item_price_id][0]'].astype(str)
df['Subscription_Item_1'] = df['subscription_items[item_price_id][1]'].astype(str) + ' _$_ ' + df['subscription_items[quantity][1]'].astype(str) + ' _$_ ' + df['subscription_items[unit_price][1]'].astype(str)
df['Subscription_Item_2'] = df['subscription_items[item_price_id][2]'].astype(str) + ' _$_ ' + df['subscription_items[quantity][2]'].astype(str) + ' _$_ '  + df['subscription_items[unit_price][2]'].astype(str)
df['Subscription_Item_3'] = df['subscription_items[item_price_id][3]'].astype(str) + ' _$_ ' + df['subscription_items[quantity][3]'].astype(str) + ' _$_ '  + df['subscription_items[unit_price][3]'].astype(str)
df['Subscription_Item_4'] = df['subscription_items[item_price_id][4]'].astype(str) + ' _$_ ' + df['subscription_items[quantity][4]'].astype(str) + ' _$_ '  + df['subscription_items[unit_price][4]'].astype(str)
df['Subscription_Item_5'] = df['subscription_items[item_price_id][5]'].astype(str) + ' _$_ ' + df['subscription_items[quantity][5]'].astype(str) + ' _$_ '  + df['subscription_items[unit_price][5]'].astype(str)
df['Subscription_Item_6'] = df['subscription_items[item_price_id][6]'].astype(str) + ' _$_ ' + df['subscription_items[quantity][6]'].astype(str) + ' _$_ '  + df['subscription_items[unit_price][6]'].astype(str)
df['Subscription_Item_7'] = df['subscription_items[item_price_id][7]'].astype(str) + ' _$_ ' + df['subscription_items[quantity][7]'].astype(str) + ' _$_ '  + df['subscription_items[unit_price][7]'].astype(str)
df['Subscription_Item_8'] = df['subscription_items[item_price_id][8]'].astype(str) + ' _$_ ' + df['subscription_items[quantity][8]'].astype(str) + ' _$_ '  + df['subscription_items[unit_price][8]'].astype(str)
df['Subscription_Item_9'] = df['subscription_items[item_price_id][9]'].astype(str) + ' _$_ ' + df['subscription_items[quantity][9]'].astype(str) + ' _$_ '  + df['subscription_items[unit_price][9]'].astype(str)



df['subscription_items'] = df.apply(lambda x: '%s' % (sorting_strings(
    [str(x['Subscription_Item_0']), str(x['Subscription_Item_1']),
     str(x['Subscription_Item_2']), str(x['Subscription_Item_3']), str(x[
         'Subscription_Item_4']), str(x['Subscription_Item_5']), str(x['Subscription_Item_6']), str(x['Subscription_Item_7']), str(x['Subscription_Item_8']), str(x['Subscription_Item_9'])
     ])), axis=1)
df1 = df[["customer[id]", "subscription[id]", "subscription[status]", "subscription[started_at]", "subscription[current_term_start]", "subscription[current_term_end]", "subscription_items", "coupon_ids[0]"]]
df1.to_excel("subscriptions_expected2_1.xlsx", index=False)
listcoltomerge = ["Subscription_Item_0", "Subscription_Item_1", "Subscription_Item_2", "Subscription_Item_3", "Subscription_Item_4", "Subscription_Item_5", "Subscription_Item_6", "Subscription_Item_7", "Subscription_Item_8", "Subscription_Item_9"]
df2 = pd.DataFrame
for col in listcoltomerge:
    df[col] = df[col].astype(str)
df['merge'] = df[listcoltomerge].values.tolist()
print(df['merge'])
df['merge1'] = df['merge'].apply(lambda x: sorting_strings(x))
print(df['merge1'])
df.to_excel("subscriptions_expected3_1.xlsx", index=False)
