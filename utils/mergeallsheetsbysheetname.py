import pandas as pd

dataFile = "/Users/cb-muneendra/Desktop/numerade_test_subs_main_25k_batch4_QA.xlsx"

if dataFile.endswith('csv'):
    df = pd.CsvFile(dataFile)
elif dataFile.endswith('xlsx'):
    df = pd.ExcelFile(dataFile)

new_df = pd.DataFrame()
for sheet in df.sheet_names:
    df1 = df.parse(sheet)
    df1["Batch"] = sheet
    new_df = new_df.append(df1, ignore_index=True)

new_df.to_excel("numerade_test_subs_DS2.xlsx", index=False)