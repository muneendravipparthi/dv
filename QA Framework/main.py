import re

import pandas as pd

from SupportingFunctions import *
from CompareData import *
from ReportGeneration import *
from RecurlyData import *
from PrevalidationChecks import *

# Get the Columns Information from Config  - >Done
# Get respective Columns from the files  --> DONE
# all the Columns data to single DataFrame  -- > DONE
# Do the cell to Cell Compare  - > Done
# generate the report  - > Done






def main():
    counter = 0
    source_df = pd.DataFrame()
    destination_df = pd.DataFrame()
    is_recurly = mapping_data['IsRecurly']
    is_recurly_ds1vsds2 = mapping_data['IsRecurlyDS1vsDS2']
    print("is_recurly = {} and is_recurly_ds1vsds2 = {}".format(is_recurly, is_recurly_ds1vsds2))
    type_of_execution = str(mapping_data['type'])
    type_of_execution = re.sub('[^A-Za-z0-9]+', '', type_of_execution)
    print(type_of_execution)
    for type_in in account_type:
        to_validate = mapping_data[type_in]
        src_key, des_key = get_key_columns(mapping_data[sort_key[counter]])

        if to_validate:
            source_files = get_files(type_source_files[counter])
            destination_files = get_files(type_destination_files[counter])
            print("source File : {}".format(source_files))
            print("destination File : {}".format(destination_files))
            source_columns, destination_columns, = get_columns(type_columns_names[counter])
            if is_recurly and is_recurly_ds1vsds2:
                print(" ------ IT's RECURLY DATA VALIDATION -----------------")
                source_df = get_recurly_source_data(type_in, source_files, source_columns)
            else:
                print(" ------ IT's NON - RECURLY and DS1vsDS2 or DS2vsDS3 DATA VALIDATION -----------------")
                source_df = read_data_from_file(source_files, source_columns)

            destination_df = read_data_from_file(destination_files, destination_columns)
            #For Eloomi
            itemids = ["1_9", "3_3", "1_18", "1_8", "3_72", "1_15", "1_6", "1_10", "1_16", "3_71", "1_19", "1_13", "6_3", "3_74", "3_65", "1_11", "3_66", "8_2", "6_1", "1_17", "5_1", "1_14", "8_1", "6_4", "8_3", "3_61", "1_20", "3_73", "3_67", "6_2", "3_40"]
            source_df_temp = source_df.copy()
            destination_df_temp = destination_df.copy()
            overalldiff = pd.DataFrame()
            for itemid in itemids:
            # itemid = '1_13'
                source_df = source_df[source_df['Item_ID'] == itemid]
                destination_df = destination_df[destination_df['Item_Id'] == itemid]

                # Filter src and dst
                source_df, destination_df = filterData(source_df, destination_df, src_key, des_key)

                # pre-validation an d formatting of the data is common
                source_df, destination_df = pre_validation_check(type_in, source_df, source_columns, destination_df, destination_columns)

                source_df = source_df.sort_values(by=src_key)
                destination_df = destination_df.sort_values(by=des_key)

                # Missing Data Capture Information

                # #Filter src and dst
                # source_df, destination_df = filterData(source_df, destination_df,src_key,des_key)
                # compare and report  code is common
                diff_df = compare_data(source_df, destination_df)  # compare
                overalldiff = overalldiff.append(diff_df)
                source_df = source_df_temp.copy()
                destination_df = destination_df_temp.copy()
                print("---------------validation_",itemid,"_completed-----------------")

            site_name, client_name = get_details('Site', 'ClientName')
            sheet_name = str(type_of_execution) + str(sheetname[counter])
            report_generation(type_in, overalldiff, site_name, client_name, module[counter], sheet_name)  # Report
            print("--------------------------------")

        else:
            print('Reading and validation is SET to FALSE')
            print("--------------------------------")

        counter += 1


if __name__ == "__main__":
    main()
