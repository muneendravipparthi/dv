from CustomerValidationChecks import *
from SubscriptionsValidationChecks import *
from InvoicesValidationChecks import *


def pre_validation_check(type_in, source_df, source_columns, destination_df, destination_columns):
    source_df.fillna('', inplace=True)
    for each_column in source_columns:
        temp_df = []
        for rows in source_df[each_column]:
            if str(rows) == '0.0' or str(rows) == '0' or str(rows).strip() == 'FALSE':
                rows = ''
            temp_df.append(str(rows).strip())
        source_df[each_column] = temp_df.copy()

    destination_df.fillna('', inplace=True)
    for each_column in destination_columns:
        temp_df = []
        for rows in destination_df[each_column]:
            if str(rows) == '0.0' or str(rows) == '0':
                rows = ''
            temp_df.append(str(rows).strip())
        destination_df[each_column] = temp_df.copy()

    if str(type_in) == 'Customers_Validation':
        source_df = customer_prevalidation_check(source_df, source_columns)
        destination_df = customer_prevalidation_check(destination_df, destination_columns)
    elif str(type_in) == 'Subscriptions_Validation':
        source_df = subscriptions_prevalidation_check(source_df, source_columns)
        destination_df = subscriptions_prevalidation_check(destination_df, destination_columns)
    elif str(type_in) == 'Invoices_Validation':
        source_df = invoices_prevalidation_check(source_df, source_columns)
        destination_df = invoices_prevalidation_check(destination_df, destination_columns)

    return source_df, destination_df
