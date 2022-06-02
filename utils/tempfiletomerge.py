import pandas as pd

file1 = "/Users/cb-muneendra/Downloads/invoices_pixellu-test_09_May_2022_13_14_21/Invoices.csv"
file2 = "/Users/cb-muneendra/Downloads/CB Extraction/DS3pixellulineitems1.xlsx"

df1 = pd.read_csv(file1)
df2 = pd.read_excel(file2)
print('invoice total rec {} ; lineitems total {}'.format(len(df1),len(df2)))
df = pd.merge(df1,df2,left_on=["Invoice Number"],right_on=["invoiceId"],how="inner")
print('invoice total rec after merge{}'.format(len(df)))
df.to_excel("newInvoiceDS3.xlsx", index=False)