import re

from tqdm import tqdm

from SupportingFunctions import *


def subscriptions_prevalidation_check(src_df, columns):
    print("Currently we are in src_subscriptions_prevalidation_check")
    str_source_columns, str_destination_columns, = get_columns("Subscriptions_Columns")
    # str_source_columns, str_destination_columns, = get_columns("String_Columns")
    str_columns = list(set(str_source_columns + str_destination_columns))
    date_source_columns, date_destination_columns, = get_columns("Date_Columns")
    date_columns = list(set(date_source_columns + date_destination_columns))
    int_source_columns, int_destination_columns, = get_columns("Int_Columns")
    int_columns = list(set(int_source_columns + int_destination_columns))
    float_source_columns, float_destination_columns, = get_columns("Float_Columns")
    float_columns = list(set(float_source_columns + float_destination_columns))
    email_source_columns, email_destination_columns, = get_columns("Email_Columns")
    email_columns = list(set(email_source_columns + email_destination_columns))
    zip_source_columns, zip_destination_columns, = get_columns("Zip_Columns")
    zip_columns = list(set(zip_source_columns + zip_destination_columns))
    dateformet = str(mapping_data['dateFormet'])
    print("Currently we are in src_customer_prevalidation_check")

    # Precondition for dates
    for col in tqdm(columns, desc='Precondition for dates'):
        if ((col in list(src_df.columns.values)) and (col in date_columns)):
            try:
                print("Converting ", col)
                src_df[col] = pd.to_datetime(src_df[col])
                src_df[col] = src_df[col].apply(
                    lambda x: pd.Timestamp(x).strftime(dateformet) if pd.isna(x) != True else None)
            except:
                print("Exception for ", col)

    # precondition for email
    for col in tqdm(columns, desc='precondition for email'):
        if (col in list(src_df.columns.values) and (col in email_columns)):
            try:
                print("Updating ", col)
                src_df[col] = src_df[col].replace('_AT_', '@', regex=True)
                src_df[col] = src_df[col].replace('@example.com', '', regex=True)
                src_df[col] = src_df[col].str.lower()
            except:
                print("Exception for ", col)

    # precondition for int
    for col in tqdm(columns, desc='precondition for int'):
        if ((col in list(src_df.columns.values)) and (col in int_columns)):
            try:
                src_df[col] = src_df[col].astype(int)
                src_df[col] = src_df[col].astype(str).apply(lambda x: x.replace('.0', ''))
                src_df = src_df.round(0)
                # src_df[col]  = src_df[col] .apply(pd.to_numeric)
                if 'phone' in col:
                    src_df[col] = re.sub('[^A-Za-z0-9]+', '', src_df[col])
            except:
                print("Exception for ", col)

    # precondition for float
    for col in tqdm(columns, desc='precondition for float'):
        if ((col in list(src_df.columns.values)) and (col in float_columns)):
            try:
                print()
                src_df = src_df.round(2)
                # src_df[col] = src_df[col].apply(lambda x: np.round(x, decimals=2))
            except:
                print("Exception for ", col)

    # precondition for zip
    for col in tqdm(columns, desc='precondition for zip'):
        if ((col in list(src_df.columns.values)) and (col in zip_columns)):
            try:
                print('zip code column type', src_df.dtypes[col])
                src_df[col] = src_df[col].values.astype(str)
                src_df[col] = src_df[col].replace('.0', '')
                src_df[col] = src_df[col].replace(' ', '')
                src_df[col] = src_df[col].str.zfill(5)
                src_df[col] = src_df[col].replace('00000', 'False')
            except:
                print("Exception for ", col)

    # precondition for string
    for col in tqdm(columns, desc='precondition for string'):
        if ((col in list(src_df.columns.values)) and (col in str_columns)):
            try:
                src_df[col] = src_df[col].str.lower()
                src_df[col] = src_df[col].astype(str).apply(lambda x: x.replace('.0', ''))
            except:
                print("Exception for ", col)

    return src_df
