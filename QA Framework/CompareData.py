import numpy as np
import pandas as pd

def compare_data(src_df, des_df):
    print('src_df rows::: ', len(src_df), '; src_df Columns::: ', len(src_df.columns))
    print('des_df rows:::', len(des_df), '; des_df Columns::: ', len(des_df.columns))
    src_df = src_df.replace(r'^\s*$', np.nan, regex=True)
    des_df = des_df.replace(r'^\s*$', np.nan, regex=True)
    des_df.equals(src_df)
    comparison_values = des_df.values == src_df.values
    rows, cols = np.where(comparison_values == False)

    for item in zip(rows, cols):
        src_df.iloc[item[0], item[1]] = '{} --> {}'.format(src_df.iloc[item[0], item[1]], des_df.iloc[item[0], item[1]])

    diff_dff = src_df.copy()
    return diff_dff

def filterData(src_df, des_df, src_key, des_key):
    srchead = list(src_df.head())
    deshead = list(des_df.head())
    try:
        src_df[src_key] = src_df[src_key].str.lower()
        des_df[des_key] = des_df[des_key].str.lower()
    except Exception as e:
        print("exception in :",e)
    srcmerged = pd.merge(src_df, des_df, how='inner', left_on=[src_key],
                      right_on=[des_key], suffixes=('', '_DROP')).filter(regex='^(?!.*_DROP)')
    desmerged = pd.merge(des_df, src_df, how='inner', left_on=[des_key],
                      right_on=[src_key], suffixes=('', '_DROP')).filter(regex='^(?!.*_DROP)')
    # srcmerged = pd.merge(src_df.drop_duplicates(subset=[src_key]), des_df.drop_duplicates(subset=[des_key]),
    #                      how='inner', left_on=[src_key],
    #                      right_on=[des_key])
    # desmerged = pd.merge(des_df.drop_duplicates(subset=[des_key]), src_df.drop_duplicates(subset=[src_key]),
    #                      how='inner', left_on=[des_key],
    #                      right_on=[src_key])
    # srcmerged = pd.merge(src_df, des_df,
    #                      how='inner', left_on=[src_key],
    #                      right_on=[des_key])
    # desmerged = pd.merge(des_df, src_df,
    #                      how='inner', left_on=[des_key],
    #                      right_on=[src_key])
    sorcedf = srcmerged[srchead]
    sorcedf.to_excel("testingsorcedf.xlsx")
    destdf = desmerged[deshead]
    destdf.to_excel("testingdestdf.xlsx")
    return sorcedf, destdf