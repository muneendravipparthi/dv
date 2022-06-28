import pandas as pd
import yaml

account_type = ['Customers_Validation', 'Subscriptions_Validation', 'Invoices_Validation']
sort_key = ['Customers_KeyColumns', 'Subscriptions_KeyColumns', 'Invoices_KeyColumns']
validation_type = ['DS1vsDS2', 'DS2vvsDS3']
type_source_files = ['Customers_SourceFile', 'Subscriptions_SourceFile', 'Invoices_SourceFile']
type_destination_files = ['Customers_DestinationFile', 'Subscriptions_DestinationFile', 'Invoices_DestinationFile']
type_columns_names = ['Customers_Columns', 'Subscriptions_Columns', 'Invoices_Columns']
comment_type = ["CUSTOMERS INFORMATION", "SUBSCRIPTIONS INFORMATION", "INVOICES INFORMATION"]
module = ["Customers Execution Report", "Subscription Execution Report", "Invoice Execution Report"]
sheetname = ["Customer_diff", "Subscription_diff", "Invoice_diff"]


# typeOfExecution = ["DS1vsDS2", "DS2vsDS3", "DS1vsDS3"]

def read_mapping_info():
    with open('config.yml') as c:
        data = yaml.load(c, Loader=yaml.FullLoader)
    read_mapping_data = data['ChargeBee Automation Configuration']
    return read_mapping_data


def read_columns_names(columns_list):
    src_columns = []
    des_columns = []
    for each in columns_list:
        column = str(each)
        column = column.replace(":", "")
        column = column.split()
        src_columns.append(column[0])
        des_columns.append(column[1])
    return src_columns, des_columns


def get_files(files):
    files_list = str(mapping_data[files])
    files_list = files_list.split(',')
    return files_list


def get_columns(columns):
    columns_name = mapping_data[columns]
    print("NO of Columns", len(columns_name))
    source_columns_names, destination_columns_names = read_columns_names(columns_name)
    return source_columns_names, destination_columns_names


def read_data_from_file(files, columns_names):
    dataframe = pd.DataFrame(columns=columns_names)
    for file in files:
        file = (str(file)).strip()
        if file.endswith('.xlsx'):
            df = pd.read_excel(file)
        elif file.endswith('.csv'):
            df = pd.read_csv(file)
        for column in columns_names:
            # print(column)
            if column in df.columns:
                # print(" {} COLUMN is present in  {} ".format(column, file))
                dataframe[column] = df[column].copy()
    return dataframe


def get_details(site, client):
    site_name = mapping_data[site]
    client_name = mapping_data[client]
    print(site_name, client_name)

    return site_name, client_name


def get_key_columns(key_columns):
    key_column = str(key_columns)
    key_column = key_column.replace(":", "")
    key_column = key_column.split()
    src_key = key_column[0]
    print("=====")
    print(key_column[0])
    print(key_column[1])
    print("=====")
    des_key = key_column[1]
    src_key = str(src_key[2::])
    des_key = str(des_key[0:len(des_key) - 2:])

    return src_key, des_key


mapping_data = read_mapping_info()
