import pandas as pd

customdelimeter = ';'
inputfile = '/Users/cb-muneendra/Downloads/export-2022-09-19T10_47_20+02_00/subscriptions/subscriptions.csv'
outputfile = 'mollieData/subscriptions.csv'

#
df = pd.read_csv(inputfile, sep = customdelimeter)

headers = list(df.head())
newheaders = {}
for ch in headers:
    newheaders[ch] = ch.replace(" ", "_")

df.rename(columns=newheaders, inplace=True)

df.to_csv(outputfile, index = False)
