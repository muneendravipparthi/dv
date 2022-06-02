from re import search
import numpy as np
import pandas as pd
import yaml
dateToFillNan = pd.Timestamp(year=1970,  month=1, day=1)


def extract_str_from_value(str_value):
    if str_value != '':
        return str_value
    else:
        return ''


def get_recurly_source_data(type_in, source_files, source_columns):
    print(" ---  get_recurly_source_data  ---")
    source_df = pd.DataFrame()
    if str(type_in) == 'Customers_Validation':
        source_df = get_customers_data(source_files, source_columns)
    elif str(type_in) == 'Subscriptions_Validation':
        source_df = get_subscriptions_data(source_files, source_columns)
    elif str(type_in) == 'Invoices_Validation':
        source_df = get_invoices_data(source_files, source_columns)

    return source_df


def read_customers_mapping_info():
    with open('RecurlyConfig.yml') as c:
        data = yaml.load(c, Loader=yaml.FullLoader)
    read_customers_mapping_data = data['Customers Automation Configuration']
    customer_account_data = read_customers_mapping_data['CustomerAddressOfAccount']
    customers_billing_data = read_customers_mapping_data['CustomerAddressOfBilling']
    return customer_account_data, customers_billing_data


def read_subscriptions_mapping_info():
    with open('RecurlyConfig.yml') as c:
        data = yaml.load(c, Loader=yaml.FullLoader)
    read_subscriptions_mapping_data = data['Subscriptions Automation Configuration']
    return read_subscriptions_mapping_data


def read_invoice_mapping_info():
    with open('RecurlyConfig.yml') as c:
        data = yaml.load(c, Loader=yaml.FullLoader)
    read_invoice_mapping_data = data['Invoices Automation Configuration']
    return read_invoice_mapping_data


def get_coupon_df(coupon_df, subscription_df):
    subscription_config_mapping = read_subscriptions_mapping_info()
    subscription_merge_key = subscription_config_mapping['Subscriptions_Merge_Key']
    customer_key = subscription_config_mapping['Customer_Account_Code']
    coupon_key = subscription_config_mapping['Coupon_Code']
    null_subscriptions = coupon_df[coupon_df.isnull().any(axis=1)]
    null_row_indexes = null_subscriptions.index.values.tolist()

    for row in null_row_indexes:
        coupon_df.drop(row, inplace=True)

    subscription_coupon_merge_df = pd.merge(subscription_df, coupon_df[[subscription_merge_key, coupon_key]], on=subscription_merge_key, how='left')
    subscription_coupon_merge_df = pd.merge(subscription_coupon_merge_df, null_subscriptions[[customer_key, coupon_key]], on=customer_key, how='left')
    subscription_coupon_merge_df['coupon'] = np.where(subscription_coupon_merge_df['coupon_code_x'].isnull(), subscription_coupon_merge_df['coupon_code_y'], subscription_coupon_merge_df['coupon_code_x'])
    subscription_coupon_merge_df.drop_duplicates(subset=subscription_merge_key, keep='first', inplace=True)

    subscription_coupon_df = pd.merge(subscription_df, subscription_coupon_merge_df[[subscription_merge_key, 'coupon']], on=subscription_merge_key, how='left')
    return subscription_coupon_df


def check_coupon_date(coupon_df, column_name):
    coupon_df[column_name] = coupon_df[column_name].fillna(dateToFillNan)
    coupon_df[column_name] = pd.to_datetime(coupon_df[column_name])
    coupon_redemption_df = coupon_df.loc[(coupon_df[column_name] > pd.Timestamp.now()) | (coupon_df[column_name] == dateToFillNan)]
    return coupon_redemption_df


def get_customers_data(source_files, source_columns):
    print("Recurly Customers..!")
    account_df = pd.DataFrame()
    billing_df = pd.DataFrame()

    for i in range(len(source_files)):
        filename = str(source_files[i]).strip()
        source_files[i] = filename
        if search("accounts", filename):
            if source_files[i].endswith('csv'):
                account_df = pd.read_csv(source_files[i], dtype='unicode')
            elif source_files[i].endswith('xlsx'):
                account_df = pd.read_excel(source_files[i])
        if search("billing", filename):
            if source_files[i].endswith('csv'):
                billing_df = pd.read_csv(source_files[i], dtype='unicode')
            elif source_files[i].endswith('xlsx'):
                billing_df = pd.read_excel(source_files[i])

    if len(source_files) >= 2:
        merge_df = pd.merge(account_df, billing_df, on='account_code', suffixes=('', '_drop'))
        merge_df.drop([col for col in merge_df.columns if 'drop' in col], axis=1, inplace=True)
        customer_account_columns, customer_billing_columns = read_customers_mapping_info()
        for i in range(len(customer_account_columns)):
            merge_df[customer_account_columns[i]].fillna(merge_df[customer_billing_columns[i]], inplace=True)
    else:
        merge_df = account_df
    return merge_df[source_columns]


def get_subscriptions_data(source_files, source_columns):
    subscription_config_mapping = read_subscriptions_mapping_info()
    subscription_df = pd.DataFrame()
    coupon_redemption_df = pd.DataFrame()
    coupon_df = pd.DataFrame()
    check_addon = subscription_config_mapping['Subscription_Mapping']['Addon']
    check_redeem = subscription_config_mapping['Subscription_Mapping']['Redemption']
    check_coupon = subscription_config_mapping['Subscription_Mapping']['Coupon']
    subscription_mapping_key = subscription_config_mapping['Subscriptions_Merge_Key']
    for file in source_files:
        file = str(file).strip()
        if "subscription" in str(file):
            subscription_df = pd.read_excel(file)
            print("subscriptions")
            subscription_key = subscription_config_mapping['Subscription_UUID']
            subscription_df.rename(columns={subscription_key: subscription_mapping_key}, inplace=True)

        if check_addon and "addon" in str(file):
            addons_df = pd.read_excel(file)
            print("addons")
            merge_key = subscription_config_mapping['Addon_Merge_Key']
            addons_df.rename(columns={merge_key: subscription_mapping_key}, inplace=True)
            no_of_addons = subscription_config_mapping['No_of_Addons']
            addon_columns = subscription_config_mapping['Addon_Columns']
            for col in addon_columns:
                addons_df[col] = addons_df[col].map(str)
                dataframe = addons_df.groupby(subscription_mapping_key)[col].agg(' '.join).reset_index()

                for i in range(no_of_addons):
                    dataframe[f'{col}_{i+1}'] = dataframe[col].str.split(' ', expand=True)[i]
                subscription_df = pd.merge(subscription_df, dataframe, on=subscription_mapping_key, how='left')

        if check_redeem and "redemptions" in str(file):
            print("redemptions")
            coupon_redemption_df = pd.read_excel(file)
            redemption_date = subscription_config_mapping['Redemption_Date']
            coupon_redemption_df = check_coupon_date(coupon_redemption_df, redemption_date)

        if check_coupon and "couponsCreated" in str(file):
            print("coupons")
            coupon_id = subscription_config_mapping['Coupon_ID']
            coupon_merge_id = subscription_config_mapping['Coupon_Merge_ID']
            coupon_df = pd.read_excel(file)
            coupon_date = subscription_config_mapping['Coupon_Date']
            coupon_df = check_coupon_date(coupon_df, coupon_date)
            coupon_df.rename(columns={coupon_id: coupon_merge_id}, inplace=True)

    if check_redeem and check_coupon:
        coupon_id = subscription_config_mapping['Coupon_Redemption_Merge_Key']
        merge_coupon = pd.merge(coupon_redemption_df, coupon_df, on=coupon_id, how='left', suffixes=('', '_drop'))
        merge_coupon.drop([col for col in merge_coupon.columns if 'drop' in col], axis=1, inplace=True)
        cols = subscription_config_mapping['Coupon_Columns']
        merge_coupon = merge_coupon[cols]
        subscription_df = get_coupon_df(merge_coupon, subscription_df)

    elif check_redeem:
        cols = subscription_config_mapping['Coupon_Columns']
        coupon_redemption_df = coupon_redemption_df[cols]
        subscription_df = get_coupon_df(coupon_redemption_df, subscription_df)

    return subscription_df[source_columns]


def get_invoices_data(source_files, source_columns):
    invoice_adjustment_combined_df = pd.DataFrame()
    invoice_config_data = read_invoice_mapping_info()

    invoice_key = invoice_config_data['Invoice_Key']
    invoice_merge_key = invoice_config_data['Invoice_Merge_Key']

    adjustment_tax_column_rename = str(invoice_config_data['Adjustment_Tax_Column_Rename']).strip()
    adjustment_tax_column_rename = str(adjustment_tax_column_rename).split(':')

    invoice_tax_column_rename = str(invoice_config_data['Invoice_Tax_Column_Rename']).strip()
    invoice_tax_column_rename = str(invoice_tax_column_rename).split(':')

    adjustment_total_column_rename = str(invoice_config_data['Adjustment_Total_Column_Rename']).strip()
    adjustment_total_column_rename = str(adjustment_total_column_rename).split(':')

    adjustment_merge_columns = invoice_config_data['Adjustment_Merge_Columns']
    payment_merge_columns = invoice_config_data['Payment_Merge_Columns']

    no_of_line_items = invoice_config_data['No_Of_LineItems']
    adjustment_columns = invoice_config_data['Invoice_Adjustment_Columns']
    adjustment_discount = invoice_config_data['Adjustment_Discount']
    adjustment_coupon_code = invoice_config_data['Adjustment_Coupon_Code']
    discounts_entity_type = invoice_config_data['Discounts_Entity_Type']
    discounts_entity_id_0 = invoice_config_data['Discounts_Entity_ID']
    discounts_amount_0 = invoice_config_data['Discounts_Amount']

    adjustment_total = pd.DataFrame()
    invoice_tax_total = pd.DataFrame()
    for file in source_files:

        if "summary" in str(file):
            invoice_adjustment_combined_df = pd.read_excel(file)
            doc_types_df = invoice_config_data['Doc_Type'][0]
            doc_type_merge_key = list(doc_types_df.items())[0][0]
            doc_type_list = list(doc_types_df.values())[0]
            doc_type_list = doc_type_list.split(',')
            invoice_adjustment_combined_df = invoice_adjustment_combined_df.loc[
                                                                                (invoice_adjustment_combined_df[doc_type_merge_key] == doc_type_list[0]) |
                                                                                (invoice_adjustment_combined_df[doc_type_merge_key] == doc_type_list[1])
                                                                                ]

        if "adjustments" in str(file):
            adjustment_df = pd.read_excel(file)
            adjustment_df.rename(columns={invoice_key: invoice_merge_key, adjustment_tax_column_rename[0].strip(): adjustment_tax_column_rename[1].strip()},
                                 inplace=True)
            cols_to_use = adjustment_df.columns.difference(invoice_adjustment_combined_df.columns)
            invoice_adjustment_combined_df.rename(columns={invoice_key: invoice_merge_key}, inplace=True)
            invoice_adjustment_combined_df = pd.merge(invoice_adjustment_combined_df, adjustment_df[cols_to_use], on=invoice_merge_key, how='left')

            for col in adjustment_columns:
                invoice_adjustment_combined_df[col] = invoice_adjustment_combined_df[col].map(str)
                dataframe = invoice_adjustment_combined_df.groupby(invoice_merge_key)[col].agg(' '.join).reset_index()
                for i in range(no_of_line_items):
                    dataframe[f'{col}_{i}'] = dataframe[col].str.split(' ', expand=True)[i]
                invoice_adjustment_combined_df = pd.merge(invoice_adjustment_combined_df, dataframe, on=invoice_merge_key, how='left')
            invoice_adjustment_combined_df.drop(['adjustment_amount_y'], axis=1, inplace=True)
            invoice_adjustment_combined_df.drop(['adjustment_product_code_y'], axis=1, inplace=True)
            invoice_adjustment_combined_df.drop(['adjustment_quantity_y'], axis=1, inplace=True)
            invoice_adjustment_combined_df[adjustment_coupon_code] = invoice_adjustment_combined_df[adjustment_coupon_code].fillna("No_Coupon")

            # Total calculation for each invoice
            adjustment_total[adjustment_total_column_rename[0].strip()] = adjustment_df[adjustment_total_column_rename[1].strip()]
            adjustment_total[adjustment_total_column_rename[0].strip()] = adjustment_total[adjustment_total_column_rename[0].strip()].fillna(0)
            adjustment_total[invoice_merge_key] = adjustment_df[invoice_merge_key]
            adjustments_data_sum_value = adjustment_total.groupby([invoice_merge_key])
            adjustment_dataframe_information = adjustments_data_sum_value.sum()
            adjustment_dataframe_information = adjustment_dataframe_information.reset_index()
            adjustment_dataframe_information[adjustment_merge_columns[0]] = invoice_adjustment_combined_df[adjustment_merge_columns[0]]
            adjustment_dataframe_information[adjustment_merge_columns[1]] = invoice_adjustment_combined_df[adjustment_merge_columns[1]]

            # Adding Payment information based on the total
            adjustment_dataframe_information[payment_merge_columns[0]] = np.where(
                (adjustment_dataframe_information[adjustment_total_column_rename[0].strip()] > 0) & (adjustment_dataframe_information[adjustment_merge_columns[1]] == 'paid'), adjustment_dataframe_information[adjustment_total_column_rename[0].strip()], 0)
            adjustment_dataframe_information[payment_merge_columns[1]] = np.where((adjustment_dataframe_information[adjustment_total_column_rename[0].strip()] > 0) & (adjustment_dataframe_information[adjustment_merge_columns[1]] == 'paid'), adjustment_dataframe_information[adjustment_merge_columns[0]], 0)
            adjustment_dataframe_information[payment_merge_columns[2]] = np.where((adjustment_dataframe_information[adjustment_total_column_rename[0].strip()] > 0) & (adjustment_dataframe_information[adjustment_merge_columns[1]] == 'paid'), "other", "")

            # tax
            invoice_tax_total[invoice_tax_column_rename[0].strip()] = adjustment_df[invoice_tax_column_rename[1].strip()]
            invoice_tax_total[invoice_tax_column_rename[0].strip()] = invoice_tax_total[invoice_tax_column_rename[0].strip()].fillna(0)
            invoice_tax_total[invoice_merge_key] = adjustment_df[invoice_merge_key]
            invoice_tax_sum = invoice_tax_total.groupby([invoice_merge_key])
            invoice_tax_information_df = invoice_tax_sum.sum()
            invoice_tax_information_df = invoice_tax_information_df.reset_index()

            invoice_adjustment_combined_df = pd.merge(invoice_adjustment_combined_df,
                                                      adjustment_dataframe_information[[invoice_merge_key, adjustment_total_column_rename[0].strip(), payment_merge_columns[0]]],
                                                      on=invoice_merge_key, how='left')
            invoice_adjustment_combined_df = pd.merge(invoice_adjustment_combined_df,
                                                      invoice_tax_information_df[[invoice_merge_key, invoice_tax_column_rename[0].strip()]],
                                                      on=invoice_merge_key, how='left')

            # Coupons
            shape = invoice_adjustment_combined_df.shape
            invoice_row_count = shape[0]
            for row in range(0, invoice_row_count):
                if invoice_adjustment_combined_df.loc[row, adjustment_coupon_code] != "No_Coupon":
                    if abs(invoice_adjustment_combined_df.loc[row, adjustment_discount]) > 0:
                        invoice_adjustment_combined_df.loc[row, adjustment_discount] = invoice_adjustment_combined_df.loc[row, adjustment_discount]
                        invoice_adjustment_combined_df.loc[row, discounts_entity_type] = "document_level_coupon"
                    else:
                        invoice_adjustment_combined_df.loc[row, adjustment_discount] = 0
                        invoice_adjustment_combined_df.loc[row, discounts_entity_type] = ""

                    invoice_adjustment_combined_df.loc[row, discounts_entity_id_0] = invoice_adjustment_combined_df.loc[row, adjustment_coupon_code]
                else:
                    adj_total = invoice_adjustment_combined_df.loc[row, f'adjustment_amount_{no_of_line_items-1}']
                    if adj_total != 0:
                        invoice_adjustment_combined_df.loc[row, discounts_entity_type] = "promotional_credits"
                    else:
                        invoice_adjustment_combined_df.loc[row, discounts_entity_type] = "document_level_coupon"
                    invoice_adjustment_combined_df.loc[row, discounts_amount_0] = adj_total
                    invoice_adjustment_combined_df.loc[row, discounts_entity_id_0] = ""

            # As duplicate rows are getting generated
            invoice_adjustment_combined_df['discounts_amounts'] = np.where(invoice_adjustment_combined_df[discounts_entity_id_0].isnull(),
                                                                           invoice_adjustment_combined_df[discounts_amount_0],
                                                                           invoice_adjustment_combined_df[adjustment_discount])
            invoice_adjustment_combined_df.drop_duplicates(subset='invoice_number_id', keep='first', inplace=True)

    # invoice_adjustment_combined_df.to_excel("Invoice_Adjustment_Merged.xlsx", index=False)
    return invoice_adjustment_combined_df[source_columns]
