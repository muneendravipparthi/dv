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

ds1File = '/Users/cb-muneendra/Downloads/bookkeeper360_test_customers_QA - stripe_raw_data.csv'
ds2File = '/Users/cb-muneendra/Downloads/customer DS2/customer DS2.csv'

# Ds1ColList = ["id", "created", "name", "email", "currency", "metadata", "name", "address.line1", "address.state",
#               "address.postal_code", "address.country"]
# Ds2ColList = ["customer[id]", "customer[created_at]", "customer[first_name]", "customer[last_name]", "customer[email]",
#               "customer[preferred_currency_code]", "customer[meta_data]", "billing_address[first_name]",
#               "billing_address[last_name]", "billing_address[line1]", "billing_address[state]", "billing_address[zip]",
#               "billing_address[country]"]
Ds1ColList = ["id", "created", "name", "email", "phone", "currency", "metadata", "name", "shipping.address.line1", "address.postal_code", "shipping.address.country"]
Ds2ColList = ["customer[id]", "customer[created_at]", "customer[first_name]", "customer[last_name]", "customer[email]", "customer[phone]", "customer[preferred_currency_code]", "customer[meta_data]", "billing_address[first_name]", "billing_address[last_name]", "billing_address[line1]", "billing_address[zip]", "billing_address[country]"]
# Ds1ColList = ["id", "created", "name", "email", "currency", "metadata", "name", "address.line1", "address.line2",
#               "address.city", "address.state", "address.postal_code", "address.country"]
# Ds2ColList = ["customer[id]", "customer[created_at]", "customer[first_name]", "customer[last_name]", "customer[email]",
#               "customer[preferred_currency_code]", "customer[meta_data]", "billing_address[first_name]",
#               "billing_address[last_name]", "billing_address[line1]", "billing_address[line2]", "billing_address[city]",
#               "billing_address[state]", "billing_address[zip]", "billing_address[country]"]
# Ds1ColList = ["id", "created", "name", "email", "phone", "currency", "metadata", "name", "address.city", "address.state", "address.postal_code", "address.country"]
# Ds2ColList = ["customer[id]", "customer[created_at]", "customer[first_name]", "customer[last_name]", "customer[email]", "customer[phone]", "customer[preferred_currency_code]", "customer[meta_data]", "billing_address[first_name]", "billing_address[last_name]", "billing_address[city]", "billing_address[state]", "billing_address[zip]", "billing_address[country]"]

# Ds1ColList = ["customer[id]", "customer[first_name]", "customer[last_name]", "customer[phone]", "customer[company]", "customer[preferred_currency_code]", "customer[email]", "payment_method[type]", "payment_method[gateway_account_id]", "payment_method[reference_id]", "customer[allow_direct_debit]", "customer[auto_collection]", "customer[taxability]", "customer[vat_number]", "customer[net_term_days]", "billing_address[first_name]", "billing_address[last_name]", "billing_address[country]"]
# Ds2ColList = ["customer.id", "customer.first_name", "customer.last_name", "customer.phone", "customer.company", "customer.preferred_currency_code", "customer.email", "customer.payment_method.type", "customer.payment_method.gateway_account_id", "customer.payment_method.reference_id", "customer.allow_direct_debit", "customer.auto_collection", "customer.taxability", "customer.vat_number", "customer.net_term_days", "customer.billing_address.first_name", "customer.billing_address.last_name", "customer.billing_address.country"]


if ds1File.endswith('csv'):
    df1 = pd.read_csv(ds1File)
elif ds1File.endswith('xlsx'):
    df1 = pd.read_excel(ds1File, sheet_name="stripe_raw_data")
# print(Ds1ColList)
# print(list(df1.columns))
df1 = df1[Ds1ColList]
try:
    df1['created'] = pd.to_datetime(df1['created'], unit='s').astype(str)
except:
    print('date convertion issue')

if ds2File.endswith('csv'):
    df2 = pd.read_csv(ds2File)
elif ds2File.endswith('xlsx'):
    df2 = pd.read_excel(ds2File, sheet_name="DS1_batch1")
df2 = df2[Ds2ColList]
df2['customer[name]'] = df2["customer[first_name]"].str.lower() + " " + df2["customer[last_name]"].str.lower()
df2['billing_address[name]'] = df2["billing_address[first_name]"].str.lower() + " " + df2[
    "billing_address[last_name]"].str.lower()
# Ds2UpdatedColList = ["customer[id]", "customer[created_at]", "customer[name]", "customer[email]",
#                      "customer[preferred_currency_code]", "customer[meta_data]", "customer[name]",
#                      "billing_address[line1]", "billing_address[state]", "billing_address[zip]",
#                      "billing_address[country]"]
Ds2UpdatedColList = ["customer[id]", "customer[created_at]", "customer[name]", "customer[email]", "customer[phone]",
                     "customer[preferred_currency_code]", "customer[meta_data]", "billing_address[name]",
                     "billing_address[line1]",
                      "billing_address[zip]", "billing_address[country]"]
# Ds2UpdatedColList = ["customer[id]", "customer[created_at]", "customer[name]", "customer[email]",
#                      "customer[preferred_currency_code]", "customer[meta_data]", "billing_address[name]",
#                      "billing_address[line1]", "billing_address[line2]", "billing_address[city]",
#                      "billing_address[state]", "billing_address[zip]", "billing_address[country]"]
df2 = df2[Ds2UpdatedColList]

# df1 = df1[df1['id'].isin(df2['customer[id]'])]
# df2 = df2[df2['customer[id]'].isin(df1['id'])]
df1.to_excel(Client + "DS1_Cus.xlsx", index=False)
print("Created DS1_Cus.xlsx")
df2.to_excel(Client + "B1_DS2_Cus.xlsx", index=False)
print("Created DS2_Cus.xlsx")

# merged = pd.merge(df1, df2, how='inner', left_on=["id"],
#                  right_on=["customer[id]"], suffixes=('_DROP', '')).filter(regex='^(?!.*_DROP)')
#
# print(merged.head())
# for c in df2.head():
#     merged.drop(c)
# merged.to_excel("testA1_DS2_Cus.xlsx", index=False)
# if [col for col in df1.columns if '1' in col] != []:
#     print('working')
