import pandas as pd
File = '/Users/cb-muneendra/Downloads/bookkeeper360_test_invoices_QA.xlsx'
xl = pd.ExcelFile(File)
df = pd.DataFrame()
print(xl.sheet_names)  # see all sheet names
# xl.parse(sheet_name)
for sheet in xl.sheet_names:
    print("started for : ", sheet)
    df_new = xl.parse(sheet)
    df_new['Batch'] = sheet
    frames = [df, df_new]
    df = pd.concat(frames)
    print(sheet," added")

df.to_excel("test.xlsx", index=False)
