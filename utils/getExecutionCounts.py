import pandas as pd

# files = ["/Users/cb-muneendra/Desktop/nlziet_Test_Subscriptions_DS2 vs DS3_Diff_B2.xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Subscriptions_DS2 vs DS3_Diff_B1.xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Subscriptions_DS2 vs DS3_Diff_B3.xlsx"]
# files = ["/Users/cb-muneendra/Desktop/nlziet_Test_Invoices_DS2 vs DS3_Diff_6(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Invoices_DS2 vs DS3_Diff_4(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Invoices_DS2 vs DS3_Diff_0(2).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Invoices_DS2 vs DS3_Diff_5(2).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Invoices_DS2 vs DS3_Diff_9(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Invoices_DS2 vs DS3_Diff_8(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Invoices_DS2 vs DS3_Diff_1(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Invoices_DS2 vs DS3_Diff_7(2).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Invoices_DS2 vs DS3_Diff_3(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Invoices_DS2 vs DS3_Diff_2(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Invoices_DS2 vs DS3_Diff_10(2).xlsx"]
# files = ["/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_23(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_3(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_11(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_7(2).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_13(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_5(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_16 (1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_19(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_18(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_2(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_0(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_4(2).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_9(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_12(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_1(2).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_17(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_21(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_8(2).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_20(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_22(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_14(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_6(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_15(1).xlsx", "/Users/cb-muneendra/Desktop/nlziet_Test_Customers_DS2 vs DS3_Diff_10(1).xlsx"]
files = ["/Users/cb-muneendra/Downloads/drive-download-20220912T085306Z-001/input_files/comma_seperated/subscriptions.csv", "/Users/cb-muneendra/Downloads/drive-download-20220912T085306Z-001/input_files/comma_seperated/products.csv", "/Users/cb-muneendra/Downloads/drive-download-20220912T085306Z-001/input_files/comma_seperated/payment-rules.csv", "/Users/cb-muneendra/Downloads/drive-download-20220912T085306Z-001/input_files/comma_seperated/invoice.csv", "/Users/cb-muneendra/Downloads/drive-download-20220912T085306Z-001/input_files/comma_seperated/invoice_lines.csv", "/Users/cb-muneendra/Downloads/drive-download-20220912T085306Z-001/input_files/comma_seperated/FinalPayments.csv", "/Users/cb-muneendra/Downloads/drive-download-20220912T085306Z-001/input_files/comma_seperated/customers.csv", "/Users/cb-muneendra/Downloads/drive-download-20220912T085306Z-001/input_files/comma_seperated/transactions.csv"]

for file in files:
    # df = pd.read_excel(file,sheet_name="Execution_Report")
    # print(df.iloc[[1, 2], [1, 2, 3]])
    # # print(df.iloc[:, [0, 1]])

    df = pd.read_csv(file);
    headers = list(df.head())
    newheaders = {}
    for ch in headers:
        newheaders[ch] = ch.replace(" ", "_")

    df.rename(columns=newheaders, inplace=True)

    df.to_csv(file, index=False)


