import pandas as pd
from pandas import ExcelWriter
from timeit import default_timer as timer
df = pd.read_csv("/Users/cb-muneendra/cb_data_validation/utils/pixellulineitemfile13.csv")

print(df.shape[0])
RowCount = df.shape[0]
FinalOutputFileName = f'converted_pixellulineitemfile13.xlsx'
writer_Difference_Records = ExcelWriter(FinalOutputFileName)
def extractStrFromValue(strValue):
    if strValue != '':
        return strValue
    else:
        return ''
previous = 0
counter  = 0
finalInvoice = pd.DataFrame()
start = timer()
for row in range(0,RowCount):
    custNum = extractStrFromValue(df.loc[row, 'Invoice Number'])
    indexer = 0
    if custNum == previous:
        print(str(custNum)  + "breaking")
        continue
    else:
        DS1RawData = df.loc[(df['Invoice Number'] == custNum)]
        inter = 0
        rowInfo = DS1RawData.shape
        for ind in DS1RawData.index:
            if inter == 0:
                entity_id = DS1RawData['Entity Id'].tolist()
                entity_type = DS1RawData['Entity Type'].tolist()
                date_From = DS1RawData['Date From'].tolist()
                date_To = DS1RawData['Date To'].tolist()
                description = DS1RawData['Description'].tolist()
                unit_Amount = DS1RawData['Unit Amount'].tolist()
                amount = DS1RawData['Amount'].tolist()

                finalInvoice.loc[indexer, 'invoiceId'] = DS1RawData.loc[ind, 'Invoice Number']
                # finalInvoice.loc[indexer, 'status'] = DS1RawData.loc[ind, 'status']
                for i in range(0,len(entity_id)):
                    finalInvoice.loc[indexer,f'entity_id[{i}]'] = entity_id[i]
                    finalInvoice.loc[indexer, f'entity_type[{i}]'] = entity_type[i]
                    finalInvoice.loc[indexer, f'date_From[{i}]'] = date_From[i]
                    finalInvoice.loc[indexer, f'date_To[{i}]'] = date_To[i]
                    finalInvoice.loc[indexer, f'description[{i}]'] = description[i]
                    finalInvoice.loc[indexer, f'unit_Amount[{i}]'] = unit_Amount[i]
                    finalInvoice.loc[indexer, f'amount[{i}]'] = amount[i]
                inter += 1
            previous = DS1RawData.loc[ind, 'Invoice Number']

        Difference_Data = finalInvoice
        DifferenceDataFrame = (pd.DataFrame(list(Difference_Data))).transpose()
        Difference_Data = Difference_Data.replace(['empty'], '')
        DifferenceDataFrame.to_excel(writer_Difference_Records, sheet_name='Invoice', startrow=0, index=False,
                                     header=False)
        Difference_Data.to_excel(writer_Difference_Records, sheet_name='Invoice', startrow=counter + 1,
                                 index=False, header=False)

        counter += 1
writer_Difference_Records.save()
print("total time: ",timer()-start)
