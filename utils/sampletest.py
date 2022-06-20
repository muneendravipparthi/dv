import pandas as pd
import  numpy as np
import datetime

transaction_df = pd.read_excel('/Users/cb-muneendra/Git/cb_data_validation/utils/transactions.xlsx')
# transaction_df = pd.read_excel('/Users/cb-muneendra/Desktop/transactions2.xlsx')

print("working on get_current_term_dates_data function")
systemdate = datetime.datetime.today() - datetime.timedelta(days=5)
today = systemdate.strftime('%Y-%m-%d')
transaction_df['system_date'] = systemdate

transaction_df['scheduled_for'] = pd.to_datetime(transaction_df['scheduled_for'], format='%Y-%m-%d')
transaction_df['system_date'] = pd.to_datetime(transaction_df['system_date'], format='%Y-%m-%d')
current_term_start = (transaction_df['scheduled_for'] <= today) & (
            (transaction_df['system_date'] - transaction_df['scheduled_for']).dt.days < 31)
# locate rows and access them using .loc() function
current_term_start = transaction_df.loc[current_term_start]
current_term_start = current_term_start.rename(columns={'scheduled_for': 'current_term_start'})


# filter rows on the basis of date
current_term_end = (transaction_df['scheduled_for'] > today) & (
            (transaction_df['scheduled_for'] - transaction_df['system_date']).dt.days < 31)
# locate rows and access them using .loc() function
current_term_end = transaction_df.loc[current_term_end]
current_term_end = current_term_end.rename(columns={'scheduled_for': 'current_term_end'})
mergedf = pd.merge(current_term_start, current_term_end, on=['customer_id', 'subscription_plan_id'], how='inner',
                   suffixes=('', '_drop'))
mergedf.drop([col for col in mergedf.columns if 'drop' in col], axis=1, inplace=True)
# tempdf = mergedf.copy()
# # duplicateDates = mergedf.duplicated(subset=['customer_id'])
# tempdf = tempdf.loc[tempdf.duplicated(), :]
# mergedf['isduplicated'] = mergedf.loc[mergedf['customer_id'].isin(tempdf['customer_id'])]
# mergedf['isduplicated'] = mergedf[['customer_id']].apply(lambda x: 'True' if x in tempdf['customer_id'] else 'False')
df = mergedf.copy()
mergedf["Duplicate"] = mergedf.duplicated(['customer_id'], keep=False)

mergedf['current_term_end'] = mergedf[['current_term_end', 'Duplicate']].apply(lambda x: x['current_term_end'] if x['Duplicate'] == False else None, axis =1)


mergedf.to_excel('1transactionmergedf1.xlsx', index=False)

def checkDiff(row):
    if row['diff'] > 0:
        return 0
    else:
        return row['current_term_end']

# df = pd.read_excel('transactionmergedf.xlsx')
df_dup = df[['customer_id','current_term_end']]
df_dup['diff'] = df_dup.groupby('customer_id')['current_term_end'].diff() / np.timedelta64(1, 'D')
df_dup['diff'] = df_dup['diff'].fillna(0)
df_dup.drop_duplicates(subset ="customer_id",keep = 'last', inplace = True)
print(df_dup.loc[(df_dup['customer_id'] ==56476290)])
df_dup['current_term_end_new'] = df_dup.apply(lambda  row: checkDiff(row),axis=1)
df = pd.merge(df,df_dup[['customer_id','current_term_end_new']],on='customer_id',how='left')
df.drop(['current_term_end'], axis = 1, inplace=True)
df.drop_duplicates(subset ="customer_id",keep = 'last', inplace = True)
df.to_excel("NewExcel.xlsx",index=False)