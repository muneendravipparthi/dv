import pandas as pd
import os
from jproperties import Properties

configs = Properties()
home_folder = os.getenv('HOME')
ROOT_DIR1 = os.path.abspath(os.curdir)
ROOT_DIR = ROOT_DIR1.replace('DS1DS2FileFilter', 'configuration.properties')
with open(ROOT_DIR, 'rb') as config_file:
    configs.load(config_file)
Client = configs.get("clientName").data

ds1File = '/Users/cb-muneendra/Downloads/bookkeeper360_test_subs_QA - stripe_raw_data.csv'
ds2File = '/Users/cb-muneendra/Downloads/subscription DS2/subscription DS2.csv'

# Ds1ColList = ["customer", "id", "created", "items.data[0].price.id", "items.data[0].quantity",
#               "items.data[0].price.unit_amount", "status", "trial_start", "trial_end", "start_date",
#               "current_period_start", "current_period_end", "collection_method", "discount.coupon.id", "metadata"]
# Ds2ColList = ["customer[id]", "subscription[id]", "subscription[created_at]", "subscription_items[item_price_id][0]",
#               "subscription_items[quantity][0]", "subscription_items[unit_price][0]", "subscription[status]",
#               "subscription[trial_start]", "subscription[trial_end]", "subscription[started_at]",
#               "subscription[current_term_start]", "subscription[current_term_end]", "subscription[auto_collection]",
#               "coupon_ids[0]", "subscription[meta_data]"]
# Ds1ColList = ["customer", "id", "created", "items.data[0].price.id", "items.data[0].quantity", "items.data[0].price.unit_amount", "status", "trial_start", "trial_end", "start_date", "current_period_start", "current_period_end", "canceled_at", "metadata"]
# Ds2ColList = ["customer[id]", "subscription[id]", "subscription[created_at]", "subscription_items[item_price_id][0]", "subscription_items[quantity][0]", "subscription_items[unit_price][0]", "subscription[status]", "subscription[trial_start]", "subscription[trial_end]", "subscription[started_at]", "subscription[current_term_start]", "subscription[current_term_end]", "subscription[cancelled_at]", "subscription[meta_data]"]
Ds1ColList = ["customer", "id", "created", "items.data[0].plan.id", "items.data[0].quantity", "items.data[0].price.unit_amount", "status", "trial_start", "trial_end", "start_date", "current_period_start", "current_period_end", "canceled_at", "collection_method", "discount.coupon.id", "metadata"]
# Ds2ColList = ["customer[id]", "subscription[id]", "subscription_items[item_price_id][0]", "subscription_items[quantity][0]", "subscription_items[unit_price][0]", "subscription[status]", "subscription[trial_start]", "subscription[trial_end]", "subscription[started_at]", "subscription[current_term_start]", "subscription[current_term_end]", "subscription[cancelled_at]", "subscription[meta_data]"]
# Ds1ColList = ["customer", "id", "created", "items.data[0].price.id", "items.data[0].quantity", "items.data[0].price.unit_amount", "status", "start_date", "current_period_start", "current_period_end", "collection_method", "discount.coupon.id", "metadata", "canceled_at"]
Ds2ColList = ["customer[id]", "subscription[id]", "subscription[created_at]", "subscription_items[item_price_id][0]", "subscription_items[quantity][0]", "subscription_items[unit_price][0]", "subscription[status]", "subscription[trial_start]", "subscription[trial_end]", "subscription[started_at]", "subscription[current_term_start]", "subscription[current_term_end]", "subscription[cancelled_at]", "subscription[auto_collection]", "coupon_ids[0]", "subscription[meta_data]"]

if ds1File.endswith('csv'):
    df1 = pd.read_csv(ds1File)
elif ds1File.endswith('xlsx'):
    df1 = pd.read_excel(ds1File)
# print(Ds1ColList)
# print(list(df1.columns))
df1 = df1[Ds1ColList]
try:
    df1['created'] = pd.to_datetime(df1['created'], errors='coerce', unit='s').astype(str)
except:
    print('created date convertion issue')
try:
    df1['start_date'] = pd.to_datetime(df1['start_date'], errors='coerce', unit='s').astype(str)
except:
    print('start_date date convertion issue')
try:
    df1['current_period_start'] = pd.to_datetime(df1['current_period_start'], errors='coerce', unit='s').astype(str)
except:
    print('current_period_start date convertion issue')
try:
    df1['current_period_end'] = pd.to_datetime(df1['current_period_end'], errors='coerce', unit='s').astype(str)
except:
    print('current_period_end date convertion issue')
try:
    df1['canceled_at'] = pd.to_datetime(df1['canceled_at'], errors='coerce', unit='s').astype(str)
except:
    print('canceled_at date convertion issue')
try:
    df1['trial_start'] = pd.to_datetime(df1['trial_start'], errors='coerce', unit='s').astype(str)
except:
    print('trial_start date convertion issue')
try:
    df1['trial_end'] = pd.to_datetime(df1['trial_end'], errors='coerce', unit='s').astype(str)
except:
    print('trial_end date convertion issue')

if ds2File.endswith('csv'):
    df2 = pd.read_csv(ds2File)
elif ds2File.endswith('xlsx'):
    df2 = pd.read_excel(ds2File)
df2 = df2[Ds2ColList]

# df1 = df1[df1.index.isin(df2.index)]
# df2 = df2[df2.index.isin(df1.index)]

df1.to_excel(Client + "_DS1_Subs.xlsx", index=False)
print("Created DS1_Subs.xlsx")
df2.to_excel(Client + "_DS2_Subs.xlsx", index=False)
print("Created DS2_Subs.xlsx")
