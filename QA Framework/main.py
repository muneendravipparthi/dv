from CompareData import *
from MollieData import *
from PrevalidationChecks import *
from RecurlyData import *
from ReportGeneration import *


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
    is_mollie = mapping_data['IsMollie']
    is_mollie_ds1vsds2 = mapping_data['IsMollieDS1vsDS2']
    print("is_mollie = {} and is_mollie_ds1vsds2 = {}".format(is_mollie, is_mollie_ds1vsds2))
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
            elif is_mollie and is_mollie_ds1vsds2:
                print(" ------ IT's MOILLE DATA VALIDATION -----------------")
                source_df = get_mollie_source_data(type_in, source_files, source_columns)
                source_df.to_excel("moille.xlsx", index=False)
            else:
                print(" ------ IT's NON - RECURLY and DS1vsDS2 or DS2vsDS3 DATA VALIDATION -----------------")
                source_df = read_data_from_file(source_files, source_columns)

            destination_df = read_data_from_file(destination_files, destination_columns)

            # # Filter src and dst
            # source_df, destination_df = filterData(source_df, destination_df, src_key, des_key)

            # pre-validation an d formatting of the data is common
            source_df, destination_df = pre_validation_check(type_in, source_df, source_columns, destination_df,
                                                             destination_columns)

            source_df = source_df.sort_values(by=src_key)
            destination_df = destination_df.sort_values(by=des_key)

            # Missing Data Capture Information

            # Filter src and dst
            source_df, destination_df = filterData(source_df, destination_df, src_key, des_key)

            if is_mollie and is_mollie_ds1vsds2:
                source_df = source_df[source_columns]
                sourceheaders = list(source_df.head())
                destheaders = list(destination_df.head())
                newheaders = {}
                for i in range(len(sourceheaders)):
                    newheaders[sourceheaders[i]] = destheaders[i]
                source_df.rename(columns=newheaders, inplace=True)

            # compare and report  code is common
            diff_df = compare_data(source_df, destination_df)  # compare
            site_name, client_name = get_details('Site', 'ClientName')
            sheet_name = str(type_of_execution) + str(sheetname[counter])
            report_generation(type_in, diff_df, site_name, client_name, module[counter], sheet_name)  # Report

            print("--------------------------------")
        else:
            print('Reading and validation is SET to FALSE')
            print("--------------------------------")

        counter += 1


if __name__ == "__main__":
    main()
