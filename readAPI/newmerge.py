import pandas as pd

df1 = pd.read_excel("/Users/cb-muneendra/Desktop/Book4.xlsx", sheet_name = "customers")
df2 = pd.read_excel("/Users/cb-muneendra/Desktop/Book4.xlsx", sheet_name = "customer-delta")
col = ["customer[id]", "customer[first_name]", "customer[last_name]", "customer[email]", "payment_method[type]", "payment_method[gateway_account_id]", "payment_method[reference_id]", "customer[auto_collection]", "customer[taxability]", "customer[preferred_currency_code]", "customer[allow_direct_debit]", "billing_address[first_name]", "billing_address[last_name]", "billing_address[email]", "billing_address[company]", "billing_address[line1]", "billing_address[line3]", "billing_address[city]", "billing_address[zip]", "billing_address[country]", "customer[cf_praktijkdata_id]", "customer[created_at]"]
df1 = df1[col]
df2 = df2[col]

df = df1.append(df2)
df.to_excel("ipractice_expected_customer.xlsx", index= False)
