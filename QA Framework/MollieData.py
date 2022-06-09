from re import search
import numpy as np
import pandas as pd
import yaml
import datetime
from datetime import datetime
from datetime import timedelta
dateToFillNan = pd.Timestamp(year=1970,  month=1, day=1)


def extract_str_from_value(str_value):
    if str_value != '':
        return str_value
    else:
        return ''


def get_mollie_source_data(type_in, source_files, source_columns):
    print(" ---  get_mollie_source_data  ---")
    source_df = pd.DataFrame()
    if str(type_in) == 'Customers_Validation':
        source_df = get_customers_data(source_files, source_columns)
    elif str(type_in) == 'Subscriptions_Validation':
        source_df = get_subscriptions_data(source_files, source_columns)
    elif str(type_in) == 'Invoices_Validation':
        source_df = get_invoices_data(source_files, source_columns)

    return source_df


def read_customers_mapping_info():
    with open('MollieConfig.yml') as c:
        data = yaml.load(c, Loader=yaml.FullLoader)
    read_customers_mapping_data = data['Customers Automation Configuration']
    customers_data = read_customers_mapping_data['Customers']
    finalPayments_data = read_customers_mapping_data['FinalPayments']
    return customers_data, finalPayments_data


def read_subscriptions_mapping_info():
    with open('MollieConfig.yml') as c:
        data = yaml.load(c, Loader=yaml.FullLoader)
    read_subscriptions_mapping_data = data['Subscriptions Automation Configuration']
    subscription_data = read_subscriptions_mapping_data['Subscriptions']
    transactions_data = read_subscriptions_mapping_data['Transactions']
    return subscription_data, transactions_data


def read_invoice_mapping_info():
    with open('MollieConfig.yml') as c:
        data = yaml.load(c, Loader=yaml.FullLoader)
    read_invoice_mapping_data = data['Invoices Automation Configuration']
    invoice_data = read_invoice_mapping_data['invoice-lines']
    transactions_data = read_invoice_mapping_data['Transactions']
    subscription_data = read_invoice_mapping_data['Subscriptions']
    return invoice_data, transactions_data, subscription_data


def get_customers_data(source_files, source_columns):
    print("Mollie Customers..!")

    for i in range(len(source_files)):
        filename = str(source_files[i]).strip()
        source_files[i] = filename
        if search("customers", filename):
            if source_files[i].endswith('csv'):
                customer_df = pd.read_csv(source_files[i], dtype='unicode')
            elif source_files[i].endswith('xlsx'):
                customer_df = pd.read_excel(source_files[i])
        if search("FinalPayments", filename):
            if source_files[i].endswith('csv'):
                payment_df = pd.read_csv(source_files[i], dtype='unicode')
            elif source_files[i].endswith('xlsx'):
                payment_df = pd.read_excel(source_files[i])

    if len(source_files) >= 2:
        merge_df = pd.merge(customer_df, payment_df, left_on='External_ID', right_on='customerId', how = 'left', suffixes=('', '_drop'))
        merge_df.drop([col for col in merge_df.columns if 'drop' in col], axis=1, inplace=True)
        customers_columns, customer_payment_columns = read_customers_mapping_info()
        # for i in range(len(customers_columns)):
        #     merge_df[customers_columns[i]].fillna(merge_df[customer_payment_columns[i]], inplace=True)
        # merge_df['mandateId'] = merge_df['External_ID'].astype(str) + "/" + merge_df['mandateId'].astype(str)
        merge_df['mandateId'] = merge_df[['External_ID', 'mandateId']].apply(lambda x: "/".join(x) if pd.isna(x.mandateId) != True else None, axis =1)
        merge_df['auto_collection'] = merge_df['mandateId'].apply(lambda x: 'OFF' if pd.isna(x) else 'ON')
        merge_df['allow_direct_debit'] = merge_df['Method'].apply(lambda x: 'TRUE' if x =='directdebit' else None)
    else:
        merge_df = customer_df
    # return merge_df[source_columns]
    return merge_df


def get_subscriptions_data(source_files, source_columns):
    print("Mollie Subscriptions..!")

    for i in range(len(source_files)):
        filename = str(source_files[i]).strip()
        source_files[i] = filename
        if search("subscriptions", filename):
            if source_files[i].endswith('csv'):
                subscription_df = pd.read_csv(source_files[i], dtype='unicode')
            elif source_files[i].endswith('xlsx'):
                subscription_df = pd.read_excel(source_files[i])
        if search("transactions", filename):
            if source_files[i].endswith('csv'):
                transaction_df = pd.read_csv(source_files[i], dtype='unicode')
            elif source_files[i].endswith('xlsx'):
                transaction_df = pd.read_excel(source_files[i])

    if len(source_files) >= 2:
        subscription_columns, transaction_columns = read_subscriptions_mapping_info()
        subscription_df = subscription_df[subscription_columns]
        temp_transaction_df = transaction_df[transaction_columns]
        transaction_df = get_currenttermdate_data(temp_transaction_df)
        merge_df = pd.merge(subscription_df, transaction_df, left_on=['Customer_ID', 'Subscriptionplan_ID'], right_on=['customer_id','subscription_plan_id'],how="left", suffixes=('', '_drop'))
        merge_df.drop([col for col in merge_df.columns if 'drop' in col], axis=1, inplace=True)

        # merge_df['mandateId'] = merge_df[['External_ID', 'mandateId']].apply(
        #     lambda x: "/".join(x) if pd.isna(x.mandateId) != True else None, axis=1)
    else:
        merge_df = subscription_df
    # return merge_df[source_columns]
    return merge_df



def get_invoices_data(source_files, source_columns):
    print("Mollie Invoices..!")

    for i in range(len(source_files)):
        filename = str(source_files[i]).strip()
        source_files[i] = filename
        if search("invoice", filename):
            if source_files[i].endswith('csv'):
                invoice_df = pd.read_csv(source_files[i], dtype='unicode')
            elif source_files[i].endswith('xlsx'):
                invoice_df = pd.read_excel(source_files[i])
        if search("transactions", filename):
            if source_files[i].endswith('csv'):
                transaction_df = pd.read_csv(source_files[i], dtype='unicode')
            elif source_files[i].endswith('xlsx'):
                transaction_df = pd.read_excel(source_files[i])
        if search("subscriptions", filename):
            if source_files[i].endswith('csv'):
                subscription_df = pd.read_csv(source_files[i], dtype='unicode')
            elif source_files[i].endswith('xlsx'):
                subscription_df = pd.read_excel(source_files[i])

    if len(source_files) >= 3:
        invoice_columns, transaction_columns, subscription_columns = read_invoice_mapping_info()
        invoice_df = invoice_df[invoice_columns]
        temp_transaction_df = transaction_df[transaction_columns]
        transaction_df = get_currenttermdate_data(temp_transaction_df)
        # transaction_df = transaction_df[transaction_columns]
        subscription_df = subscription_df[subscription_columns]

        merge_df = pd.merge(subscription_df, transaction_df, left_on=['Customer_ID', 'Subscriptionplan_ID'],
                            right_on=['customer_id', 'subscription_plan_id'], how="left", suffixes=('', '_drop'))
        merge_df.drop([col for col in merge_df.columns if 'drop' in col], axis=1, inplace=True)

        # merge_df = pd.merge(subscription_df, transaction_df, left_on='Customer_ID', right_on='customer_id',
        #                     suffixes=('', '_drop'))
        # merge_df.drop([col for col in merge_df.columns if 'drop' in col], axis=1, inplace=True)

        merge_df = pd.merge(invoice_df, merge_df, left_on=['Customer_ID','Product_ID'], right_on=['Customer_ID','Subscriptionplan_ID'], how="left",
                            suffixes=('', '_drop'))
        merge_df.drop([col for col in merge_df.columns if 'drop' in col], axis=1, inplace=True)
        merge_df.drop("customer_id", axis=1, inplace=True)

        # merge_df['mandateId'] = merge_df[['External_ID', 'mandateId']].apply(
        #     lambda x: "/".join(x) if pd.isna(x.mandateId) != True else None, axis=1)
    else:
        merge_df = invoice_df
    # return merge_df[source_columns]
    return merge_df

def get_currenttermdate_data(transaction_df):
    print("working on currentterm dates")
    # today = datetime.datetime.today().strftime('%Y-%m-%d')
    # Previous_Date = (datetime.datetime.today() - datetime.timedelta(days=31)).strftime('%Y-%m-%d')
    # Next_Date = (datetime.datetime.today() + datetime.timedelta(days=31)).strftime('%Y-%m-%d')
    today = '2022-05-29'
    today = datetime.strptime(today,'%Y-%m-%d')
    Next_Date = today + timedelta(days=31)
    Previous_Date = today - timedelta(days=31)
    print(today, Previous_Date, Next_Date)

    transaction_df['scheduled_for'] = pd.to_datetime(transaction_df['scheduled_for'], format='%Y-%m-%d')
    # filter rows on the basis of date
    current_term_start = (transaction_df['scheduled_for'] >= Previous_Date) & (transaction_df['scheduled_for'] <= today)
    # locate rows and access them using .loc() function
    current_term_start = transaction_df.loc[current_term_start]
    current_term_start = current_term_start.rename(columns={'scheduled_for': 'current_term_start'})

    # filter rows on the basis of date
    current_term_end = (transaction_df['scheduled_for'] >= today) & (transaction_df['scheduled_for'] <= Next_Date)
    # locate rows and access them using .loc() function
    current_term_end = transaction_df.loc[current_term_end]
    current_term_end = current_term_end.rename(columns={'scheduled_for': 'current_term_end'})
    mergedf = pd.merge(current_term_start, current_term_end, on=['customer_id','subscription_plan_id'], how='inner',
                     suffixes=('', '_drop'))
    mergedf.drop([col for col in mergedf.columns if 'drop' in col], axis=1, inplace=True)
    return mergedf