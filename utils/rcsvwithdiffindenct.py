import pandas as pd

# inputfiles ="/Users/cb-muneendra/Downloads/input 2/subscriptions/subscriptions.csv"
# inputfiles = "/Users/cb-muneendra/Downloads/input 2/transactions/transactions.csv"
inputfiles = "/Users/cb-muneendra/Downloads/input 2/invoices/invoice-lines.csv"
# outputfile = 'customers.xlsx'
# outputfile = 'subscriptions.xlsx'
# outputfile = 'transactions.xlsx'
outputfile = 'invoice.xlsx'

if inputfiles.endswith('csv'):
    df = pd.read_csv(inputfiles, sep = ';', engine = 'python')
    headers = list(df.head())
    newheaders = {}
    for ch in headers:
        newheaders[ch] = ch.replace(" ", "_")
    df.rename(columns=newheaders, inplace=True)
    df.to_excel(outputfile, index=False)
else:
    print("Input file is not csv")